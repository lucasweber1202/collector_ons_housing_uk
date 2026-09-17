# collector_ons_housing_uk

Standalone collector focused on the official ONS Price Index of Private Rents (PIPR) successor dataset. It stores index, annual inflation and published rent level for the UK, Great Britain, four nations and nine English regions: 45 series and 6,114 monthly observations from 2015-01 through 2026-08.

The collector does not splice the structurally different legacy IPHRP method onto PIPR. Latest rows use the official timestamp; historical values in the current workbook are `first_seen` unless archived evidence exists.

## Install and run (PowerShell)

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install .
Copy-Item .env.example .env
pytest -q
python main.py
```

Set `COLLECTOR_DB_URL` and allow `ons.gov.uk`. Databricks is optional via `.[databricks]`. Source smoke: `python -c "from scripts.extract import collect; x=collect(); print(len(x.catalog), len(x.observations))"`.

See [METHODOLOGY.md](METHODOLOGY.md) and [POINT_IN_TIME.md](POINT_IN_TIME.md).
