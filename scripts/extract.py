"""ONS Price Index of Private Rents successor series (PIPR), January 2015 onward."""

from __future__ import annotations

import hashlib
import io
import re
from dataclasses import dataclass
from datetime import UTC, date, datetime
from typing import Any
from urllib.parse import urljoin

import httpx
import openpyxl

from scripts.config import MAX_DOWNLOAD_BYTES, REQUEST_TIMEOUT, USER_AGENT
from scripts.snapshots import Snapshot, build_snapshot
from scripts.time_series import Observation

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
