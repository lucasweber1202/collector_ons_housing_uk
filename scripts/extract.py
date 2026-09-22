"""ONS Price Index of Private Rents successor series (PIPR), January 2015 onward."""

from __future__ import annotations

import hashlib
import io
import logging
import math
import re
from dataclasses import dataclass
from datetime import UTC, date, datetime
from typing import Any
from urllib.parse import urljoin

import httpx
import openpyxl

from scripts.config import (
    MAX_DOWNLOAD_BYTES,
    MAX_STALE_MONTHS,
    MIN_HISTORY_YEARS,
    REQUEST_TIMEOUT,
    USER_AGENT,
)
from scripts.snapshots import Snapshot, build_snapshot
from scripts.time_series import Observation

logger = logging.getLogger(__name__)


# -- series_id contract (GUIDELINES.md 4) ---------------------------------
# series_id is uppercase, underscore-separated and ordered coarse -> fine. The
# pair below is the canonical public surface: parse splits an id into its
# components, build rejoins them, and build(*parse(sid)) == sid for every id
# this collector emits. Only economic identity is encoded -- never a delivery
# provider or any other detail of how the value reached us.


def parse_series_id(series_id: str) -> tuple[str, ...]:
    """Split a series_id into its underscore-delimited components.

    Raises ValueError on anything this collector would not have produced:
    lowercase, empty components, or an id with no structure at all.
    """
    if not series_id or series_id != series_id.upper():
        raise ValueError(f"series_id must be uppercase: {series_id!r}")
    components = tuple(series_id.split("_"))
    if any(not component for component in components):
        raise ValueError(f"series_id has an empty component: {series_id!r}")
    return components


def build_series_id(*components: str) -> str:
    """Rejoin the tuple parse_series_id returned into the original id."""
    if not components:
        raise ValueError("series_id needs at least one component")
    if any(not component or component != component.upper() for component in components):
        raise ValueError(f"invalid series_id components: {components!r}")
    return "_".join(components)


# -- 5.1 usable-series filtering ------------------------------------------


@dataclass(frozen=True)
class UsabilityReport:
    """What the filter removed, for logging and for tests to assert on."""

    kept: tuple[str, ...]
    stale: tuple[str, ...]
    short_history: tuple[str, ...]
    empty: tuple[str, ...]

    @property
    def dropped(self) -> tuple[str, ...]:
        return tuple(sorted(set(self.stale) | set(self.short_history) | set(self.empty)))


def _months_between(earlier: date, later: date) -> int:
    """Whole months from ``earlier`` to ``later``, day-of-month aware."""
    months = (later.year - earlier.year) * 12 + (later.month - earlier.month)
    if later.day < earlier.day:
        months -= 1
    return months


def _is_valid(value: Any) -> bool:
    """A real observation: present, numeric and finite."""
    if value is None:
        return False
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        return False
    return math.isfinite(numeric)


def assess_series(
    reference_dates: list[date],
    today: date,
    max_stale_months: int = MAX_STALE_MONTHS,
    min_history_years: float = MIN_HISTORY_YEARS,
) -> str:
    """Classify one series from the reference dates of its valid observations.

    Returns ``"keep"``, ``"empty"``, ``"stale"`` or ``"short_history"``.
    Recency is judged at the period end and over non-null values only: a source
    that keeps listing a discontinued series with empty recent cells must not
    look live because of those blanks.
    """
    if not reference_dates:
        return "empty"
    first, last = min(reference_dates), max(reference_dates)
    if _months_between(last, today) > max_stale_months:
        return "stale"
    if _months_between(first, last) < round(min_history_years * 12):
        return "short_history"
    return "keep"


def filter_usable_series(
    observations: list[Any],
    catalog: dict[str, dict[str, Any]],
    today: date,
    max_stale_months: int = MAX_STALE_MONTHS,
    min_history_years: float = MIN_HISTORY_YEARS,
) -> tuple[list[Any], dict[str, dict[str, Any]], UsabilityReport]:
    """Drop obsolete and history-less series before anything is persisted.

    Runs after parsing and before the time_series / metadata upsert, so the
    standardized tables never carry a dead or stub series, and prunes the
    catalog alongside the observations so metadata can never describe a series
    the database does not hold (GUIDELINES.md 5.1).
    """
    valid_dates: dict[str, list[date]] = {}
    for observation in observations:
        if _is_valid(observation.value):
            valid_dates.setdefault(observation.series_id, []).append(observation.reference_date)

    verdicts: dict[str, str] = {}
    for series_id in set(catalog) | {o.series_id for o in observations}:
        verdicts[series_id] = assess_series(
            valid_dates.get(series_id, []), today, max_stale_months, min_history_years
        )

    keep = {series_id for series_id, verdict in verdicts.items() if verdict == "keep"}
    report = UsabilityReport(
        kept=tuple(sorted(keep)),
        stale=tuple(sorted(s for s, v in verdicts.items() if v == "stale")),
        short_history=tuple(sorted(s for s, v in verdicts.items() if v == "short_history")),
        empty=tuple(sorted(s for s, v in verdicts.items() if v == "empty")),
    )

    if report.dropped:
        logger.info(
            "Usable-series filter: kept %d, dropped %d "
            "(stale=%d short_history=%d empty=%d; max_stale_months=%d min_history_years=%s)",
            len(report.kept),
            len(report.dropped),
            len(report.stale),
            len(report.short_history),
            len(report.empty),
            max_stale_months,
            min_history_years,
        )
        for series_id in report.stale:
            logger.info(
                "Dropped %s: last valid observation older than %d months",
                series_id,
                max_stale_months,
            )
        for series_id in report.short_history:
            logger.info(
                "Dropped %s: valid history shorter than %s years", series_id, min_history_years
            )
        for series_id in report.empty:
            logger.info("Dropped %s: no valid observations", series_id)
    else:
        logger.info("Usable-series filter: all %d series usable", len(report.kept))

    kept_observations = [o for o in observations if o.series_id in keep]
    kept_catalog = {sid: fields for sid, fields in catalog.items() if sid in keep}
    return kept_observations, kept_catalog, report


LANDING = "https://www.ons.gov.uk/economy/inflationandpriceindices/datasets/priceindexofprivaterentsukmonthlypricestatistics"
AREA_CODES = {
    "K02000001",
    "K03000001",
    "E92000001",
    "W92000004",
    "S92000003",
    "N92000002",
    *(f"E1200000{i}" for i in range(1, 10)),
}
MEASURES = {5: ("INDEX", "index"), 7: ("ANNUAL_CHANGE", "percent"), 8: ("RENTAL_PRICE", "currency")}


@dataclass(frozen=True)
class ExtractedData:
    observations: list[Observation]
    snapshots: list[Snapshot]
    catalog: dict[str, dict[str, Any]]
    releases: list[datetime]
    availability_by_key: dict[tuple[str, date], tuple[datetime, str, date | None]]
    min_lag_days: int = 0
    max_lag_days: int = 0
    inferred_lag_days: int | None = None


def parse_xlsx(
    body: bytes, snapshot_id: str, url: str, collected: datetime
) -> tuple[
    list[Observation],
    dict[str, dict[str, Any]],
    datetime,
    dict[tuple[str, date], tuple[datetime, str, date | None]],
]:
    book = openpyxl.load_workbook(io.BytesIO(body), data_only=True, read_only=True)
    if "Cover sheet" not in book.sheetnames or "Table 1" not in book.sheetnames:
        raise ValueError("ONS PIPR workbook sheets drifted")
    cover = " ".join(str(book["Cover sheet"].cell(r, 1).value or "") for r in range(1, 10))
    m = re.search(
        r"published at (\d{1,2}:\d{2})(?:am|pm)? on (\d{1,2} \w+ \d{4})", cover, re.IGNORECASE
    )
    if not m:
        raise ValueError("ONS PIPR official publication timestamp missing")
    released = datetime.strptime(f"{m.group(2)} {m.group(1)}", "%d %B %Y %H:%M").replace(tzinfo=UTC)
    sheet = book["Table 1"]
    expected = (
        "Time period",
        "Area code",
        "Area name",
        "Region or country name",
        "Index",
        "Monthly change",
        "Annual change",
        "Rental price",
    )
    if tuple(sheet.cell(3, c).value for c in range(1, 9)) != expected:
        raise ValueError("ONS PIPR required columns drifted")
    observations = []
    catalog = {}
    availability = {}
    keys = set()
    for row in sheet.iter_rows(min_row=4, values_only=True):
        reference, code, name = row[0], str(row[1]), str(row[2])
        if code not in AREA_CODES or not isinstance(reference, datetime):
            continue
        for col, (measure, unit) in MEASURES.items():
            raw = row[col - 1]
            if not isinstance(raw, (int, float)):
                continue
            series_id = f"ONS_PIPR_{code}_{measure}_JAN2023"
            key = (series_id, reference.date().replace(day=1))
            if key in keys:
                raise ValueError(f"Duplicate ONS PIPR key {key}")
            keys.add(key)
            observations.append(Observation(series_id, key[1], float(raw), snapshot_id))
            catalog[series_id] = {
                "source_id": "ons_private_rents_pipr",
                "name": f"{name} private rents {measure.lower().replace('_', ' ')}",
                "description": "Raw PIPR value; methodology starts January 2015 and index base January 2023 is retained in series_id. No automatic concatenation to IPHRP.",
                "frequency": "monthly",
                "unit": unit,
                "eco_group": "inflation",
                "source_url": url,
                "last_publish_date": released.date(),
            }
    latest = max(o.reference_date for o in observations)
    for o in observations:
        current = o.reference_date == latest
        availability[(o.series_id, o.reference_date)] = (
            released if current else collected,
            "official_timestamp" if current else "first_seen",
            released.date() if current else None,
        )
    if len(catalog) != 45 or len(observations) < 5000:
        raise ValueError("ONS PIPR coverage unexpectedly changed")
    return observations, catalog, released, availability


def collect() -> ExtractedData:
    fetched = datetime.now(UTC)
    with httpx.Client(
        timeout=REQUEST_TIMEOUT, headers={"User-Agent": USER_AGENT}, follow_redirects=True
    ) as client:
        page = client.get(LANDING)
        page.raise_for_status()
        links = re.findall(
            r'href=["\']([^"\']*file\?uri=[^"\']+\.xlsx)["\']', page.text, re.IGNORECASE
        )
        if not links:
            raise ValueError("No ONS PIPR workbook discovered")
        url = urljoin(LANDING, links[0])
        response = client.get(url)
        response.raise_for_status()
    body = response.content
    if not body or len(body) > MAX_DOWNLOAD_BYTES:
        raise ValueError(f"Invalid ONS PIPR artifact size {len(body)}")
    digest = hashlib.sha256(body).hexdigest()
    obs, catalog, release, availability = parse_xlsx(body, digest, url, fetched)
    snapshot = build_snapshot(
        "ons_private_rents_pipr",
        url,
        "ons_pipr.xlsx",
        body,
        digest,
        response.headers.get("etag"),
        response.headers.get("last-modified"),
        fetched,
        release.date(),
    )
    return ExtractedData(obs, [snapshot], catalog, [release], availability)
