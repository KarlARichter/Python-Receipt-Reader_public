"""
hyperlink.py

One-off helper: turns the plain RenamedFileName text (column L) in each region
sheet into a clickable relative hyperlink to the image in ProcessedReceiptsRenamed.
Run it once to fix rows that were written before hyperlinks were added. Safe to
re-run; rows that are already hyperlinks are skipped.

Uses the same ROOT / EXCEL_FILENAME as process_receipts.py, so set those there.
Close the workbook in Excel before running.
"""

import os
import re

import openpyxl
from openpyxl.styles import Font

# match your main script
from process_receipts import (
    ROOT, EXCEL_FILENAME, SUBDIR_RENAMED,
    REGION_FOLDER_PATTERN, FIRST_DATA_ROW,
)

RENAMED_COL = 12  # column L


def main():
    path = os.path.join(ROOT, EXCEL_FILENAME)
    if not os.path.exists(path):
        print("Workbook not found at {}.".format(path))
        return

    wb = openpyxl.load_workbook(path)
    changed = 0

    for sheet in wb.sheetnames:
        if not re.match(REGION_FOLDER_PATTERN, sheet):
            continue  # skip CAD_Tax_Rates and anything non-region
        ws = wb[sheet]
        for r in range(FIRST_DATA_ROW, ws.max_row + 1):
            cell = ws.cell(row=r, column=RENAMED_COL)
            val = cell.value
            if not val:
                continue
            if isinstance(val, str) and val.startswith("=HYPERLINK("):
                continue  # already done
            name = str(val)
            rel = "{}\\{}\\{}".format(sheet, SUBDIR_RENAMED, name)
            esc = lambda s: str(s).replace('"', '""')
            cell.value = '=HYPERLINK("{}","{}")'.format(esc(rel), esc(name))
            cell.font = Font(color="0563C1", underline="single")
            changed += 1

    wb.save(path)
    print("Done. {} filename(s) turned into hyperlinks.".format(changed))


if __name__ == "__main__":
    main()