"""Build the Google Sheets template used by the Make scenario.

The classification logic lives in SHEET FORMULAS (helper columns H..R), so the Make
scenario stays tiny: it only fetches rows whose status changed, sends one grouped
message, and writes `last_alert_status` back.

Input columns  (A..G): data a human maintains; Make writes only column G.
Helper columns (H..R): formulas, never edited by hand.

Usage:
    python scripts/build_template.py        # writes data/inventory_make_template.xlsx
Import the .xlsx into Google Sheets (File > Import) and keep the first sheet's name: inventory
"""
from __future__ import annotations

import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
CSV_PATH = ROOT / "data" / "inventory_sample.csv"
OUT_PATH = ROOT / "data" / "inventory_make_template.xlsx"

FONT = "Arial"
INPUT_COLUMNS = ["sku", "product_name", "category", "stock", "reorder_point", "supplier", "last_alert_status"]

# Helper columns: (header, formula template). {r} is replaced by the row number.
# Thresholds: OUT = 0, CRITICAL <= 50% of reorder point, LOW <= reorder point.
HELPERS = [
    ("stock_num",         '=IFERROR(VALUE(TRIM(D{r}&"")),"")'),
    ("reorder_num",       '=IFERROR(VALUE(TRIM(E{r}&"")),"")'),
    ("status",            '=IF(OR(H{r}="",H{r}<0),"DATA_ERROR",IF(OR(I{r}="",I{r}<=0),"DATA_ERROR",'
                          'IF(H{r}=0,"OUT",IF(H{r}<=I{r}*0.5,"CRITICAL",IF(H{r}<=I{r},"LOW","OK")))))'),
    ("issue",             '=IF(J{r}<>"DATA_ERROR","",IF(OR(H{r}="",H{r}<0),"invalid stock: \'"&D{r}&"\'",'
                          '"invalid reorder_point: \'"&E{r}&"\'"))'),
    ("severity",          '=IF(J{r}="OUT",3,IF(J{r}="CRITICAL",2,IF(J{r}="LOW",1,0)))'),
    ("previous_status",   '=IF(TRIM(G{r}&"")="","OK",UPPER(TRIM(G{r})))'),
    ("previous_severity", '=IF(M{r}="OUT",3,IF(M{r}="CRITICAL",2,IF(M{r}="LOW",1,0)))'),
    ("should_notify",     '=IF(OR(L{r}>N{r},AND(J{r}="DATA_ERROR",M{r}<>"DATA_ERROR")),1,0)'),
    ("needs_update",      '=IF(J{r}<>M{r},1,0)'),
    ("sort_priority",     '=IF(J{r}="OUT",4,IF(J{r}="CRITICAL",3,IF(J{r}="LOW",2,IF(J{r}="DATA_ERROR",1,0))))'),
    ("alert_line",        '=IF(J{r}="OK","",IF(J{r}="DATA_ERROR","\u26a0\ufe0f "&A{r}&" "&B{r}&" \u2014 "&K{r},'
                          'IF(J{r}="OUT","\U0001F534",IF(J{r}="CRITICAL","\U0001F7E0","\U0001F7E1"))'
                          '&" "&J{r}&": "&A{r}&" "&B{r}&" \u2014 stock "&H{r}&" (reorder at "&I{r}&") \u00b7 "&F{r}))'),
]
# Letters used inside the formulas above (keep in sync with the order of HELPERS):
# H stock_num, I reorder_num, J status, K issue, L severity, M previous_status,
# N previous_severity, O should_notify, P needs_update, Q sort_priority,
# R alert_line


def read_rows(path: Path = CSV_PATH) -> list[list]:
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        next(reader)  # header
        return [row for row in reader]


def build_workbook(rows: list[list], out_path: Path | str = OUT_PATH, with_help_sheet: bool = True) -> None:
    """`rows`: list of 7-item lists in INPUT_COLUMNS order. Values may be str, int, float or None.
    Strings that look like numbers are written as numbers (as Google Sheets does on CSV import)
    unless the caller passes them as non-str types explicitly."""
    out_path = Path(out_path)
    wb = Workbook()
    ws = wb.active
    ws.title = "inventory"

    input_fill = PatternFill("solid", start_color="1F3864")
    helper_fill = PatternFill("solid", start_color="7F7F7F")
    headers = INPUT_COLUMNS + [h for h, _ in HELPERS]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
        cell.fill = input_fill if c <= len(INPUT_COLUMNS) else helper_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r, row in enumerate(rows, start=2):
        for c, value in enumerate(row, start=1):
            if value == "":
                value = None  # empty cell, like a blank CSV field
            ws.cell(row=r, column=c, value=value).font = Font(name=FONT, size=10)
        for offset, (_, formula) in enumerate(HELPERS):
            col = len(INPUT_COLUMNS) + 1 + offset
            ws.cell(row=r, column=col, value=formula.format(r=r)).font = Font(name=FONT, size=10, color="404040")

    last = len(rows) + 1
    status_col = get_column_letter(len(INPUT_COLUMNS) + 3)  # J
    rng = f"{status_col}2:{status_col}{last}"
    for text, color in (("OUT", "F8CBAD"), ("CRITICAL", "FCE4D6"), ("LOW", "FFF2CC"),
                        ("DATA_ERROR", "D9D9D9"), ("OK", "E2EFDA")):
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="equal", formula=[f'"{text}"'], fill=PatternFill("solid", bgColor=color)))

    ws.freeze_panes = "C2"
    widths = {"A": 10, "B": 24, "C": 12, "D": 8, "E": 13, "F": 17, "G": 18, "H": 10, "I": 12,
              "J": 12, "K": 28, "L": 9, "M": 16, "N": 17, "O": 13, "P": 13, "Q": 13, "R": 70}
    for letter, w in widths.items():
        ws.column_dimensions[letter].width = w
    ws.row_dimensions[1].height = 32

    if with_help_sheet:
        h = wb.create_sheet("how_to_use")
        notes = [
            "Inventory alerts: how this sheet works",
            "",
            "Blue columns (A-G) are data. Grey columns (H-R) are formulas: do not edit them.",
            "Make writes ONLY column G (last_alert_status), always with the value of the status column.",
            "Leave G empty when you start; empty and OK mean the same thing.",
            "",
            "status: OUT = stock 0 | CRITICAL = stock <= 50% of reorder_point | LOW = stock <= reorder_point",
            "        OK = above reorder_point | DATA_ERROR = stock/reorder_point blank, text or invalid",
            "should_notify = 1 only when severity got WORSE than last_alert_status (or a new data error).",
            "needs_update = 1 when last_alert_status must be rewritten (includes silent recoveries).",
            "alert_line is the ready-made text line Make sends to Telegram.",
            "",
            "Keep the first sheet named: inventory",
            "Keep the header names exactly as they are: the Make scenario refers to them.",
        ]
        for i, line in enumerate(notes, start=1):
            c = h.cell(row=i, column=1, value=line)
            c.font = Font(name=FONT, size=12 if i == 1 else 10, bold=(i == 1))
        h.column_dimensions["A"].width = 110

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def main() -> None:
    rows = read_rows()
    # numeric-looking inputs become numbers; blanks and text such as "n/a" stay as they are
    def convert(v: str):
        try:
            return int(v)
        except ValueError:
            return v
    typed = [[r[0], r[1], r[2], convert(r[3]), convert(r[4]), r[5], r[6]] for r in rows]
    build_workbook(typed)
    print(f"Wrote {OUT_PATH} ({len(rows)} products)")


if __name__ == "__main__":
    main()
