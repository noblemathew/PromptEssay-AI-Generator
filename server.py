
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
                "website_sheet": "National",         # tab to read in website file

                "old_data_start_row": 8,          # first data row in old file
                "old_first_col": 1,               # first data column (A = 1)

                "website_data_start_row": 3,         # first data row in website file
                "website_first_col": 1,              # first data column in website

                # LAST column that may be overwritten. Everything to the
                # right of this (calculated columns) is left untouched.
                # Accepts a letter ("O") or a number (15). None = no limit.
                "old_last_col": "O",

                # Optional: after pasting, drag the calculated columns down
                # so they cover the new row count. Set to the first formula
                # column, e.g. "P". Leave as None to not touch them at all.
                "formula_fill_from_col": None,
                "formula_fill_to_col": None,

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

# Excel must be VISIBLE for cell selection to work. Set to False only if
# you don't care about the file opening with A1 selected.
EXCEL_VISIBLE = True

# The date in the file name is found automatically and rolled forward to
# the next occurrence of this weekday. Monday=0, Tue=1, Wed=2, Thu=3, Fri=4.
FRIDAY_WEEKDAY = 4

FILENAME_DATE_FORMAT = "%m%d%y"

# Blank rows between the last data row and the disclaimer block.
BLANK_ROWS_BEFORE_DISCLAIMER = 1

# =======================================================================


def find_date_in_name(stem):
    """Locate a 6-digit MMDDYY date anywhere in the file name stem.

    Returns (date, start_index, end_index). Prefers a trailing date; if the
    name has several 6-digit groups, the LAST valid one wins.
    """
    candidates = list(re.finditer(r"\d{6}", stem))
    if not candidates:
        raise ValueError(
            f"No 6-digit MMDDYY date found in '{stem}' "
            "(expected something like REPORT_071726)"
        )

    for m in reversed(candidates):
        try:
            d = datetime.strptime(m.group(0), FILENAME_DATE_FORMAT)
        except ValueError:
            continue  # e.g. 999999 - not a real date, keep looking
        return d, m.start(), m.end()

    raise ValueError(
        f"Found 6-digit group(s) in '{stem}' but none is a valid "
        f"{FILENAME_DATE_FORMAT} date."
    )


def next_friday(from_date):
    """The next Friday strictly AFTER from_date.

    If from_date is itself a Friday, this returns the Friday 7 days later,
    so a weekly file always moves forward one week.
    """
    # Monday = 0 ... Friday = 4
    days_ahead = (FRIDAY_WEEKDAY - from_date.weekday()) % 7
    if days_ahead == 0:
        days_ahead = 7
    return from_date + timedelta(days=days_ahead)


def build_new_filename(old_path):
    """Read the date out of the file name and roll it to the next Friday."""
    folder, filename = os.path.split(os.path.abspath(old_path))
    stem, ext = os.path.splitext(filename)

    old_date, i, j = find_date_in_name(stem)
    new_date = next_friday(old_date)

    print(f"    file date {old_date:%d-%b-%Y} ({old_date:%A}) "
          f"-> {new_date:%d-%b-%Y} ({new_date:%A})")

    new_stem = stem[:i] + new_date.strftime(FILENAME_DATE_FORMAT) + stem[j:]
    return os.path.join(folder, new_stem + ext)


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


def col_letter(n):
    """1 -> A, 15 -> O."""
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def col_number(letters):
    """A -> 1, O -> 15. Also accepts an int and passes it through."""
    if isinstance(letters, int):
        return letters
    n = 0
    for ch in str(letters).strip().upper():
        if not ch.isalpha():
            raise ValueError(f"Bad column reference: {letters!r}")
        n = n * 26 + (ord(ch) - 64)
    return n


def go_to_a1(sheet):
    """Leave the tab with A1 selected and scrolled to the top-left.

    Range.Select only works on the ACTIVE sheet of a VISIBLE Excel window,
    so this activates the sheet first and falls back through gentler
    methods if the environment still refuses.
    """
    try:
        sheet.api.Activate()
    except Exception:
        pass

    # Preferred: Goto scrolls and selects in one step
    try:
        sheet.book.app.api.Goto(sheet.api.Range("A1"), True)
        return
    except Exception:
        pass

    # Fallback: plain Select
    try:
        sheet.api.Range("A1").Select()
        return
    except Exception:
        pass

    # Last resort: just scroll the view, no selection change
    try:
        win = sheet.book.app.api.ActiveWindow
        win.ScrollRow = 1
        win.ScrollColumn = 1
    except Exception:
        print("    NOTE: could not reset the view to A1 (harmless).")


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

    # --- Work out the LAST column we are allowed to write to ---
    # Anything beyond old_last_col (e.g. calculated columns) is never
    # cleared and never written to.
    old_last_col = cfg.get("old_last_col")
    if old_last_col is None:
        write_last_col = old_first_col + n_cols - 1
    else:
        if isinstance(old_last_col, str):
            old_last_col = col_number(old_last_col)
        write_last_col = old_last_col
        allowed = write_last_col - old_first_col + 1
        if n_cols > allowed:
            print(f"    NOTE: website has {n_cols} columns but only {allowed} "
                  f"(up to {col_letter(write_last_col)}) may be written - "
                  "extra website columns ignored.")
            n_cols = allowed
        elif n_cols < allowed:
            print(f"    WARNING: website has only {n_cols} columns but config "
                  f"allows up to {col_letter(write_last_col)} - "
                  f"columns {col_letter(old_first_col + n_cols)}"
                  f"-{col_letter(write_last_col)} will be left blank.")
            write_last_col = old_first_col + n_cols - 1

    # --- Clear the old values in the data area, then paste the new ones ---
    ws.range(
        (data_start, old_first_col), (new_last_data_row, write_last_col)
    ).clear_contents()

    # Trim/pad each row to exactly the columns we are writing
    width = write_last_col - old_first_col + 1
    padded = [list(r)[:width] + [None] * (width - len(r[:width]))
              for r in new_rows]
    ws.range((data_start, old_first_col)).value = padded

    print(f"    wrote {new_count} rows x {width} cols "
          f"({col_letter(old_first_col)}{data_start}:"
          f"{col_letter(write_last_col)}{new_last_data_row})")

    # --- Extend the calculated columns to cover the new row count ---
    fill_from = cfg.get("formula_fill_from_col")
    if fill_from is not None and new_count > 1:
        if isinstance(fill_from, str):
            fill_from = col_number(fill_from)
        fill_to = cfg.get("formula_fill_to_col")
        if fill_to is None:
            fill_to = last_used_col(ws)
        elif isinstance(fill_to, str):
            fill_to = col_number(fill_to)

        if fill_to >= fill_from:
            src = ws.range((data_start, fill_from), (data_start, fill_to))
            dst = ws.range((data_start, fill_from),
                           (new_last_data_row, fill_to))
            src.api.AutoFill(dst.api, 0)  # 0 = xlFillDefault
            print(f"    filled formulas {col_letter(fill_from)}-"
                  f"{col_letter(fill_to)} down to row {new_last_data_row}")
    print(f"    disclaimer now starts at row {new_disc_start} "
          f"({BLANK_ROWS_BEFORE_DISCLAIMER} blank row after data)")

    # --- Put the cursor back at A1 ---
    go_to_a1(ws)


def process_report(report):
    old_file = os.path.abspath(report["old_file"])
    website_file = os.path.abspath(report["website_file"])

    print(f"\n=== {os.path.basename(old_file)} ===")

    new_file = build_new_filename(old_file)
    if os.path.exists(new_file):
        raise ValueError(f"Output already exists, not overwriting: {new_file}")

    app = xw.App(visible=EXCEL_VISIBLE)
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
