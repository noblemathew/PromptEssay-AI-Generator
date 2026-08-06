import os
import re
from datetime import datetime, timedelta

import xlwings as xw

# =============================== CONFIG ================================

REPORTS = [
    {
        "old_file": r"REPORT_1_071726.xlsx",
        "website_file": r"website_report_1.xlsx",
        "sheets": [
            {
                "old_sheet": "Weekly Report",     # tab to update in old file
                "website_sheet": "",         # tab to read in website file

                "old_data_start_row": 8,          # first data row in old file
                "old_first_col": 1,               # first data column (A = 1)

                "website_data_start_row": 3,         # first data row in website file
                "website_first_col": 1,              # first data column in website

                # A distinctive phrase from the disclaimer block.
                "disclaimer_marker_text": "REPLACE_WITH_DISCLAIMER_SNIPPET",

                # If the marker phrase is NOT on the first line of the
                # disclaimer block, set how many block rows sit above it.
                "disclaimer_rows_above_marker": 0,
            },
            # Second tab for this report - copy the block above if needed.
        ],
    },
    # More reports: same shape.
]

DAYS_TO_ADVANCE = 7
FILENAME_DATE_FORMAT = "%m%d%y"

# Blank rows between the last data row and the disclaimer block.
BLANK_ROWS_BEFORE_DISCLAIMER = 1

# =======================================================================


def build_new_filename(old_path, days=DAYS_TO_ADVANCE):
    """Advance the trailing MMDDYY date in the file name by `days`."""
    folder, filename = os.path.split(os.path.abspath(old_path))
    stem, ext = os.path.splitext(filename)

    match = re.search(r"(\d{6})$", stem)
    if not match:
        raise ValueError(
            f"No trailing 6-digit MMDDYY date in '{filename}' "
            "(expected e.g. REPORT_071726.xlsx)"
        )

    old_date = datetime.strptime(match.group(1), FILENAME_DATE_FORMAT)
    new_date = (old_date + timedelta(days=days)).strftime(FILENAME_DATE_FORMAT)
    return os.path.join(folder, stem[: match.start(1)] + new_date + ext)


def last_used_row(sheet):
    return sheet.api.UsedRange.Row + sheet.api.UsedRange.Rows.Count - 1


def last_used_col(sheet):
    return sheet.api.UsedRange.Column + sheet.api.UsedRange.Columns.Count - 1


def find_marker_row(sheet, start_row, marker_text):
    """Scan rows from start_row down for a cell containing marker_text."""
    end_row = last_used_row(sheet)
    end_col = last_used_col(sheet)
    needle = marker_text.strip().lower()

    for r in range(start_row, end_row + 1):
        values = sheet.range((r, 1), (r, end_col)).value
        if not isinstance(values, list):
            values = [values]
        for v in values:
            if v is not None and needle in str(v).lower():
                return r
    return None


def read_website_block(sheet, data_start_row, first_col):
    """Read the website data as a list of rows, stopping at the first blank row."""
    end_row = last_used_row(sheet)
    end_col = last_used_col(sheet)

    block = sheet.range((data_start_row, first_col), (end_row, end_col)).value
    if block is None:
        return []
    if not isinstance(block, list):
        block = [[block]]
    if block and not isinstance(block[0], list):
        block = [block]

    rows = []
    for row in block:
        if all(v is None or str(v).strip() == "" for v in row):
            break  # stop at the first fully blank row
        rows.append(row)
    return rows


def process_sheet(old_wb, website_wb, cfg):
    old_sheet_name = cfg["old_sheet"]
    website_sheet_name = cfg["website_sheet"]

    old_names = [s.name for s in old_wb.sheets]
    website_names = [s.name for s in website_wb.sheets]
    if old_sheet_name not in old_names:
        raise ValueError(f"Tab '{old_sheet_name}' not in old file. Tabs: {old_names}")
    if website_sheet_name not in website_names:
        raise ValueError(f"Tab '{website_sheet_name}' not in website file. Tabs: {website_names}")

    ws = old_wb.sheets[old_sheet_name]
    website_ws = website_wb.sheets[website_sheet_name]

    print(f"  Tab '{old_sheet_name}'  <-  website tab '{website_sheet_name}'")

    data_start = cfg["old_data_start_row"]
    old_first_col = cfg["old_first_col"]

    # --- New data ---
    new_rows = read_website_block(
        website_ws, cfg["website_data_start_row"], cfg["website_first_col"]
    )
    if not new_rows:
        raise ValueError(
            f"No data rows found in website tab '{website_sheet_name}' - "
            "stopping so the old data isn't wiped for nothing."
        )
    new_count = len(new_rows)
    n_cols = max(len(r) for r in new_rows)

    # --- Locate the disclaimer block ---
    marker_row = find_marker_row(ws, data_start, cfg["disclaimer_marker_text"])
    if marker_row is None:
        raise ValueError(
            f"Disclaimer marker not found in tab '{old_sheet_name}' - "
            "check disclaimer_marker_text."
        )

    disc_start = marker_row - cfg.get("disclaimer_rows_above_marker", 0)
    disc_end = last_used_row(ws)
    disc_height = disc_end - disc_start + 1

    if disc_start <= data_start:
        raise ValueError(
            f"Disclaimer appears to start at row {disc_start}, which is at or "
            f"above the data start row {data_start}. Check the config."
        )

    # Old data occupied everything up to the blank gap before the disclaimer
    old_last_data_row = disc_start - 1 - BLANK_ROWS_BEFORE_DISCLAIMER
    old_count = old_last_data_row - data_start + 1

    print(f"    old data rows: {old_count}  |  new data rows: {new_count}")
    print(f"    disclaimer block: rows {disc_start}-{disc_end} ({disc_height} rows)")

    # --- Shift the disclaimer block by inserting / deleting rows above it ---
    delta = new_count - old_count
    if delta > 0:
        # Insert blank rows just above the disclaimer block, pushing it down
        ws.api.Rows(f"{disc_start}:{disc_start + delta - 1}").Insert()
        print(f"    inserted {delta} row(s) - disclaimer moved down")
    elif delta < 0:
        # Remove surplus rows from just above the disclaimer block
        n = -delta
        ws.api.Rows(f"{disc_start - n}:{disc_start - 1}").Delete()
        print(f"    deleted {n} row(s) - disclaimer moved up")
    else:
        print("    row count unchanged - disclaimer stays put")

    new_last_data_row = data_start + new_count - 1
    new_disc_start = new_last_data_row + 1 + BLANK_ROWS_BEFORE_DISCLAIMER

    # --- Clear the old values in the data area, then paste the new ones ---
    clear_end_col = max(
        old_first_col + n_cols - 1,
        last_used_col(ws),
    )
    ws.range(
        (data_start, old_first_col), (new_last_data_row, clear_end_col)
    ).clear_contents()

    # Normalise row widths so xlwings writes a clean rectangle
    padded = [list(r) + [None] * (n_cols - len(r)) for r in new_rows]
    ws.range((data_start, old_first_col)).value = padded

    print(f"    wrote {new_count} rows (rows {data_start}-{new_last_data_row})")
    print(f"    disclaimer now starts at row {new_disc_start} "
          f"({BLANK_ROWS_BEFORE_DISCLAIMER} blank row after data)")

    # --- Put the cursor back at A1 ---
    ws.activate()
    ws.range("A1").select()


def process_report(report):
    old_file = os.path.abspath(report["old_file"])
    website_file = os.path.abspath(report["website_file"])

    print(f"\n=== {os.path.basename(old_file)} ===")

    new_file = build_new_filename(old_file)
    if os.path.exists(new_file):
        raise ValueError(f"Output already exists, not overwriting: {new_file}")

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    old_wb = website_wb = None
    try:
        old_wb = app.books.open(old_file)
        website_wb = app.books.open(website_file)

        for sheet_cfg in report["sheets"]:
            process_sheet(old_wb, website_wb, sheet_cfg)

        # Land on the first updated tab at A1
        old_wb.sheets[report["sheets"][0]["old_sheet"]].activate()

        old_wb.save(new_file)
        print(f"  Saved as: {os.path.basename(new_file)}")
    finally:
        if website_wb is not None:
            website_wb.close()
        if old_wb is not None:
            old_wb.close()
        app.quit()


def main():
    ok = failed = 0
    for report in REPORTS:
        try:
            process_report(report)
            ok += 1
        except Exception as e:
            failed += 1
            print(f"  ERROR ({report.get('old_file', '?')}): {e}")

    print(f"\nFinished. {ok} succeeded, {failed} failed.")
    input("Press Enter to close...")


if __name__ == "__main__":
    main()
