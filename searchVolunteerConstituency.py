import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook


INPUT_FILE = Path("CAREkakis Volunteers Performance Matrix_R.xlsx")
SEARCH_CONSTITUENCY = "Tampines West"
OUTPUT_FILE = Path("tampines_west_events.xlsx")


TARGET_COLUMNS = [
    "Date",
    "Day",
    "Time",
    "Venue",
    "Program Stage",
    "Program Name",
    "Remarks",
    "Staff",
    "Trainer",
    "Registered",
    "GRL Attended",
    "Public Attended CA",
    "Total Attendance",
    "Source Sheet",
    "Original Header",
]


VENUE_MAP = {
    "BL": "Boon Lay",
    "KG": "Kampong Glam",
    "KPG": "Kampong Glam",
    "BBE": "Bukit Batok East",
    "TB": "Tampines",
    "TW": "Tampines West",
    "VAD": "Varies / To confirm",
}


def normalise_text(value):
    if value is None:
        return ""
    return str(value).strip()


def is_target_header_row(row_values):
    row_norm = [normalise_text(v).lower() for v in row_values]
    return "s/n" in row_norm and "name" in row_norm and "constituency" in row_norm


def find_header_row(ws, max_scan_rows=10):
    for r in range(1, min(ws.max_row, max_scan_rows) + 1):
        values = [ws.cell(r, c).value for c in range(1, min(ws.max_column, 20) + 1)]
        if is_target_header_row(values):
            return r
    return None


def is_attendance_mark(value):
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return value != 0
    text = str(value).strip()
    if text == "":
        return False
    if text in {"-", "0", "N", "No", "NO", "n", "no"}:
        return False
    return True


def header_looks_like_event(header_text):
    if not header_text:
        return False

    text = str(header_text).strip()

    # Ignore summary / admin columns
    ignore_phrases = [
        "total cc",
        "total ck",
        "total csg",
        "total cc + csg",
        "cumulative total",
        "no. of touchpoints",
        "signed up as volunteer",
        "training completed",
        "statement of attainment",
        "certification",
        "t-shirt",
        "email",
        "handphone",
        "constituency",
        "name",
        "s/n",
        "checklist",
        "javin update",
        "updated as of",
        "legend",
    ]

    lowered = text.lower()
    if any(phrase in lowered for phrase in ignore_phrases):
        return False

    # Event headers usually contain a date-like pattern
    date_patterns = [
        r"\b\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}\b",   # 22 Feb 2025
        r"\b\d{1,2}\s+[A-Za-z]{3,9}\b",           # 11 April
        r"\b\d{1,2}\s+[A-Za-z]{3,9}\s*:",         # 12 Feb:
        r"\bpre-?ns\b",
        r"\bpurple parade\b",
    ]

    return any(re.search(p, text, flags=re.IGNORECASE) for p in date_patterns)


def parse_event_header(header_text, source_sheet):
    """
    Best-effort parser.
    Because the workbook stores event info inside column headers, not all fields
    are available. Missing fields are left blank.
    """
    original = normalise_text(header_text).replace("\n", " ").strip()
    compact = re.sub(r"\s+", " ", original)

    date_value = None
    day_value = ""
    time_value = ""
    venue_value = ""
    program_stage = ""
    program_name = compact
    remarks = ""

    # 1. Try to parse full date with year
    full_date_patterns = [
        r"(?P<date>\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4})",
        r"(?P<date>\d{1,2}\s+[A-Za-z]{3,9}\s+\d{2})",
    ]

    for pattern in full_date_patterns:
        m = re.search(pattern, compact, flags=re.IGNORECASE)
        if m:
            raw_date = m.group("date")
            for fmt in ("%d %b %Y", "%d %B %Y", "%d %b %y", "%d %B %y"):
                try:
                    date_value = datetime.strptime(raw_date, fmt).date()
                    break
                except ValueError:
                    continue
            if date_value:
                break

    # 2. If no year in header, infer from sheet name
    if date_value is None:
        inferred_year = None
        year_match = re.search(r"(20\d{2}|24|23)", source_sheet)
        if year_match:
            y = year_match.group(1)
            inferred_year = int("20" + y) if len(y) == 2 else int(y)

        partial_date_match = re.search(
            r"(?P<date>\d{1,2}\s+[A-Za-z]{3,9})(?:\s*:|\b)",
            compact,
            flags=re.IGNORECASE,
        )
        if partial_date_match and inferred_year:
            raw_date = f"{partial_date_match.group('date')} {inferred_year}"
            for fmt in ("%d %b %Y", "%d %B %Y"):
                try:
                    date_value = datetime.strptime(raw_date, fmt).date()
                    break
                except ValueError:
                    continue

    if date_value:
        day_value = date_value.strftime("%A")

    # 3. Try to extract venue code from patterns like "12 Feb: KPG: Needs Assess"
    venue_match = re.search(
        r"^\s*\d{1,2}\s+[A-Za-z]{3,9}(?:\s+\d{4})?\s*:\s*([A-Za-z]{2,4})\s*:",
        compact,
        flags=re.IGNORECASE,
    )
    if venue_match:
        venue_code = venue_match.group(1).upper()
        venue_value = VENUE_MAP.get(venue_code, venue_code)

    # 4. Remove leading date part from program name
    program_name = re.sub(
        r"^\s*\d{1,2}\s+[A-Za-z]{3,9}(?:\s+\d{4})?\s*[:\-]?\s*",
        "",
        compact,
        flags=re.IGNORECASE,
    ).strip()

    # If venue code appears at start, remove it from program name
    program_name = re.sub(r"^[A-Za-z]{2,4}\s*:\s*", "", program_name).strip()

    # Special case
    if "pre-ns forum" in compact.lower():
        program_name = compact

    return {
        "Date": date_value,
        "Day": day_value,
        "Time": time_value,
        "Venue": venue_value,
        "Program Stage": program_stage,
        "Program Name": program_name,
        "Remarks": remarks,
        "Staff": "",
        "Trainer": "",
        "Registered": 0,
        "GRL Attended": "",
        "Public Attended CA": "",
        "Total Attendance": 0,
        "Source Sheet": source_sheet,
        "Original Header": compact,
    }


def extract_events_for_constituency(input_file, constituency):
    wb = load_workbook(input_file, data_only=True)
    results = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        header_row = find_header_row(ws)
        if header_row is None:
            continue

        headers = {
            c: normalise_text(ws.cell(header_row, c).value)
            for c in range(1, ws.max_column + 1)
        }

        # Find required columns by header text
        constituency_col = None
        for c, h in headers.items():
            if h.lower() == "constituency":
                constituency_col = c
                break

        if constituency_col is None:
            continue

        # Build event column list
        event_cols = [
            c for c, h in headers.items()
            if header_looks_like_event(h)
        ]

        if not event_cols:
            continue

        # Find matching volunteer rows
        matching_rows = []
        for r in range(header_row + 1, ws.max_row + 1):
            cell_value = normalise_text(ws.cell(r, constituency_col).value)
            if cell_value.lower() == constituency.strip().lower():
                matching_rows.append(r)

        if not matching_rows:
            continue

        # Count attendance marks per event column
        for c in event_cols:
            header_text = headers[c]
            attended_count = 0

            for r in matching_rows:
                value = ws.cell(r, c).value
                if is_attendance_mark(value):
                    attended_count += 1

            # Keep only events where at least one matching volunteer has a mark
            if attended_count > 0:
                event_row = parse_event_header(header_text, sheet_name)
                event_row["Registered"] = attended_count
                event_row["Total Attendance"] = attended_count
                results.append(event_row)

    return pd.DataFrame(results, columns=TARGET_COLUMNS)


def main():
    df = extract_events_for_constituency(INPUT_FILE, SEARCH_CONSTITUENCY)

    if df.empty:
        print(f'No marked events found for constituency: "{SEARCH_CONSTITUENCY}"')
        print("This may mean either:")
        print("1. there are no attendance marks for that constituency in this workbook, or")
        print("2. the workbook is tracking volunteer home constituency, not event constituency.")
        return

    # Sort by date where possible
    df["_sort_date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.sort_values(by=["_sort_date", "Program Name"], na_position="last").drop(columns=["_sort_date"])

    print(df.to_string(index=False))

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Extracted Events")

    print(f"\nSaved to: {OUTPUT_FILE.resolve()}")


if __name__ == "__main__":
    main()