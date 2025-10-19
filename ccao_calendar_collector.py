"""
CCAO Assessment Calendar Collector — Local Excel Version (no Google APIs)
Collects key dates from the Cook County Assessor's Assessment Calendar and saves as an .xlsx file.

Usage:
    python ccao_calendar_collector.py
Optional:
    Place a "tri schedule.csv" file (same structure as before) next to this script to include triennial info.
"""
import os
import re
import sys
import json
import time
import errno
from datetime import datetime, timedelta

import requests
import pandas as pd
from bs4 import BeautifulSoup

# Optional Excel formatting
try:
    import openpyxl  # for column autosize and freezing header
except Exception:
    openpyxl = None

URL = "https://www.cookcountyassessor.com/assessment-calendar-and-deadlines"
TRI_CSV = os.path.join(os.path.dirname(__file__), "tri schedule.csv")

# -----------------------
# Helpers
# -----------------------
def format_date(date_str: str) -> str:
    """Return 'Weekday, Month Dth, YYYY' from 'M/D/YYYY'. Fallback to 'TBD'."""
    try:
        dt = datetime.strptime(date_str, "%m/%d/%Y")
        day = dt.day
        suffix = "th" if 11 <= day <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
        return dt.strftime(f"%A, %B {day}{suffix}, %Y")
    except Exception:
        return "TBD"

def _get_time(row, field_class):
    el = row.select_one(f".{field_class} time")
    return format_date(el.get_text(strip=True)) if el else "TBD"

def _get_bor_range(row):
    # Prefer two <time> tags (range), support one <time>, else text fallback
    times = row.select(".field--name-field-board-of-review-appeal-dat time")
    if not times:
        times = row.select(".field--name-field-board-of-review-appeal-dates time")
    if len(times) >= 2:
        start = format_date(times[0].get_text(strip=True))
        end   = format_date(times[1].get_text(strip=True))
        return f"{start} - {end}"
    if len(times) == 1:
        return format_date(times[0].get_text(strip=True))

    fld = row.select_one(".field--name-field-board-of-review-appeal-dat") or row.select_one(".field--name-field-board-of-review-appeal-dates")
    if not fld:
        return "TBD"
    txt = fld.get_text(" ", strip=True)
    month = r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)"
    long = rf"{month}\s+\d{{1,2}},\s*\d{{4}}"
    short = r"\d{1,2}/\d{1,2}/\d{4}"
    pat = re.compile(rf"({long}|{short})")
    found = pat.findall(txt)
    flat = []
    for d in found:
        if isinstance(d, (list, tuple)):
            picks = [x for x in d if x]
            flat.append(picks[-1] if picks else "")
        else:
            flat.append(d)
    flat = [x for x in flat if x]
    if len(flat) >= 2:
        return f"{format_date(flat[0])} - {format_date(flat[1])}"
    if len(flat) == 1:
        return format_date(flat[0])
    return "TBD"

def _format_short(d) -> str:
    return f"{d.month}/{d.day}/{d.year}"

def _parse_one_date_token(tok):
    s = tok.strip()
    s = re.sub(r'^[A-Za-z]+,\s+', '', s)  # drop weekday
    s = re.sub(r'(\d{1,2})(st|nd|rd|th)', r'\1', s)  # drop ordinal
    for fmt in ("%B %d, %Y", "%b %d, %Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    m = re.search(r'(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|'
                  r'Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)'
                  r'\s+\d{1,2},\s*\d{4}', s)
    if m:
        for fmt in ("%B %d, %Y", "%b %d, %Y"):
            try:
                return datetime.strptime(m.group(0), fmt).date()
            except Exception:
                pass
    m2 = re.search(r'\d{1,2}/\d{1,2}/\d{4}', s)
    if m2:
        try:
            return datetime.strptime(m2.group(0), "%m/%d/%Y").date()
        except Exception:
            return None
    return None

def split_bor_dates_to_open_close(bor_text):
    if not bor_text or not str(bor_text).strip() or str(bor_text).strip().upper() == "TBD":
        return "", ""
    txt = str(bor_text).replace("–", "-").replace("—", "-")
    parts = [p for p in txt.split("-") if p.strip()]
    if len(parts) >= 2:
        d_open = _parse_one_date_token(parts[0])
        d_close = _parse_one_date_token(parts[1])
        return (_format_short(d_open) if d_open else "",
                _format_short(d_close) if d_close else "")
    else:
        d = _parse_one_date_token(txt)
        ds = _format_short(d) if d else ""
        return (ds, ds) if ds else ("", "")

def calc_bor_evidence_deadline(close_str):
    if not close_str:
        return ""
    try:
        d = datetime.strptime(close_str, "%m/%d/%Y").date()
        return _format_short(d + timedelta(days=10))
    except Exception:
        return ""

def determine_tri_label(years, current):
    if current in years:
        return "Yes"
    elif (current - 1) in years:
        return "No - 2nd Year of Tri"
    elif (current - 2) in years:
        return "No - 3rd Year of Tri"
    else:
        return "No"

def load_triennial():
    """Load tri schedule if present. Return DataFrame or None."""
    if not os.path.exists(TRI_CSV):
        return None
    tri = pd.read_csv(TRI_CSV)
    tri["Years"] = tri["Years"].astype(str)
    tri["Years"] = tri["Years"].apply(lambda x: [int(y.strip()) for y in x.split(",") if y.strip().isdigit()])
    current_year = datetime.now().year
    tri["Re-assessment Year"] = tri["Years"].apply(lambda y: determine_tri_label(y, current_year))
    tri = tri.drop(columns=["Years", "Re-assessment 1", "Re-assessment 2", "Re-assessment 4"], errors="ignore")
    tri = tri.rename(columns={"Re-assessment 3": "Next Triennial Year"})
    return tri

# -----------------------
# Gather calendar entries
# -----------------------
def gather_calendar():
    resp = requests.get(URL, timeout=30)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    data = []
    for row in soup.select("div.views-row"):
        try:
            title_link = row.select_one(".views-field-title a")
            if not title_link:
                continue
            township = title_link.get_text(strip=True)
            mailed_date   = _get_time(row, "field--name-field-reassessment-notice-date")
            deadline_date = _get_time(row, "field--name-field-last-file-date")
            a_roll_certified = _get_time(row, "field--name-field-date-a-roll-certified")
            a_roll_published = _get_time(row, "field--name-field-date-a-roll-published")
            bor_dates        = _get_bor_range(row)
            bor_open, bor_close = split_bor_dates_to_open_close(bor_dates)
            bor_evidence_deadline = calc_bor_evidence_deadline(bor_close)
            bor_open_fmt = format_date(bor_open) if bor_open else ""
            bor_close_fmt = format_date(bor_close) if bor_close else ""
            bor_evidence_deadline_fmt = format_date(bor_evidence_deadline) if bor_evidence_deadline else ""

            published = "Yes" if mailed_date != "TBD" and deadline_date != "TBD" else "No"
            timestamp = format_date(datetime.now().strftime("%m/%d/%Y"))

            data.append({
                "Township": township,
                "Reassessment Notices Mailed": mailed_date,
                "Assessor Appeal Deadline": deadline_date,
                "Date A-Roll Certified": a_roll_certified,
                "Date A-Roll Published": a_roll_published,
                "BOR Open For Filing Complaint": bor_open_fmt,
                "BOR Closed For Filing Complaint": bor_close_fmt,
                "BOR Evidence Submission Deadline": bor_evidence_deadline_fmt,
                "Published?": published,
                "Last Updated": timestamp,
            })
        except Exception as e:
            # Keep going even if a row has issues
            print(f"Skipping row due to error: {e}")
    df = pd.DataFrame(data)
    tri = load_triennial()
    if tri is not None:
        df = df.merge(tri, on="Township", how="left")
    # Remove duplicates: if township appears more than once, keep only published
    if "Township" in df.columns:
        dupe_counts = df["Township"].value_counts()
        dupe_townships = dupe_counts[dupe_counts > 1].index.tolist()
        df_dupes = df[df["Township"].isin(dupe_townships)]
        cleaned_dupes = df_dupes[df_dupes["Published?"].str.lower() == "yes"]
        df = df[~df["Township"].isin(dupe_townships)]
        df = pd.concat([df, cleaned_dupes], ignore_index=True)
    # Column order
    desired_cols = [
        "Township",
        "Reassessment Notices Mailed",
        "Assessor Appeal Deadline",
        "Date A-Roll Certified",
        "Date A-Roll Published",
        "BOR Open For Filing Complaint",
        "BOR Closed For Filing Complaint",
        "BOR Evidence Submission Deadline",
        "Published?",
        "Last Updated",
        "Next Triennial Year",
        "Re-assessment Year",
    ]
    # Keep only those that exist
    columns = [c for c in desired_cols if c in df.columns]
    df = df[columns]
    return df

# -----------------------
# Save to Excel
# -----------------------
def save_excel(df: pd.DataFrame, out_path: str):
    df.to_excel(out_path, index=False)
    if openpyxl:
        try:
            wb = openpyxl.load_workbook(out_path)
            ws = wb.active
            # Freeze first row
            ws.freeze_panes = "A2"
            # Autosize columns
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    try:
                        val = str(cell.value) if cell.value is not None else ""
                    except Exception:
                        val = ""
                    if len(val) > max_len:
                        max_len = len(val)
                ws.column_dimensions[col_letter].width = min(max_len + 2, 60)
            wb.save(out_path)
        except Exception as e:
            print(f"Excel formatting skipped: {e}")

def main():
    print("Collecting calendar entries...")
    df = gather_calendar()
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M")
    out_name = f"CCAO_Calendar_{ts}.xlsx"
    out_path = os.path.join(os.getcwd(), out_name)
    save_excel(df, out_path)
    print(f"Saved: {out_path}")

if __name__ == "__main__":
    main()