# CCAO Calendar Collector

This small utility collects important dates from the Cook County Assessor's Assessment Calendar and writes a clean Excel spreadsheet for easy reference. No Google credentials or external APIs are required.

## Quick start

```bash
git clone https://github.com/<your-username>/ccao-calendar-scraper.git
cd ccao-calendar-scraper
python -m venv .venv
# Activate the venv: source .venv/bin/activate  (Windows: .venv\Scripts\activate)
pip install -r requirements.txt
python ccao_calendar_collector.py
```

The output file will be saved in the current directory with a name like `CCAO_Calendar_2025-10-18_15-42.xlsx`.

## Optional: Triennial schedule
If you maintain a `tri schedule.csv` file with a `Township` column and a `Years` column (comma-separated years), place it next to the script and it will be merged into the result.

## Notes
- Designed for local use and redistribution.
- Keep dependencies minimal: requests, beautifulsoup4, pandas, openpyxl.
- The script is defensive against missing fields and will continue to run when encountering unexpected markup.