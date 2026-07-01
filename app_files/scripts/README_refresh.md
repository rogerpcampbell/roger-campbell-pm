# Refresh workflow

Use the in-app uploader for the normal weekly workflow.

For CLI refresh:

```bash
BOP_REPORTS_DIR=reports BOP_BUNDLE_PATH=data/weekly_data_bundle.json python scripts/extract_weekly_data.py
```

The parser keeps all parsed watchlist history and uses year + week labels so uploaded 2024, 2025, 2026, and future reports can be trended without mixing identical week numbers across different years.

The parser extracts BOP-level HSE KPIs, overall schedule data, waypoints, risk summary, and scope-linked watchlist items for Rail, Roads, and Ponds. It uses `pdftotext` when available and falls back to `pypdf` on Windows.
