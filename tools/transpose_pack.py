#!/usr/bin/env python3
"""Flip every data sheet of a pack onto its side. Specification_v1.md §6.2.

A submission arrives however the cedent keeps it, and plenty of them run years across the
top rather than down the side. §6.2 says the tool reads both, and `01` and `02` have
carried transposed twins since the beginning — but *only* those two. Whether every other
dataset survives the flip was, until this tool, an assumption.

So this reads a finished row-wise pack and writes the same figures transposed: labels
down a column, records across. Nothing is retyped, which is the point — any difference in
the output is the tool's, not the fixture's.

Run from the repository root::

    python tools/transpose_pack.py Intake_FireCatFull_v1.xlsx
"""

from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter as col_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from sheets import blue, body_f, green, grey, head_f, inert_f, marker_f  # noqa: E402
from sheets import put, sub_f, title_f, widths, yellow  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]

FIRST_RECORD_COL = 4          # D — B holds the sheet's own words, C the declared label
LABEL_COL, SOURCE_COL = "C", "B"


def _markers(ws):
    """Column A, row by row. Structure is separated from the rest on the way out."""
    out = []
    for row in range(1, ws.max_row + 1):
        text = ws.cell(row=row, column=1).value
        if text not in (None, ""):
            out.append((row, str(text).strip()))
    return out


def _blocks(ws):
    """One entry per ``Header_i``: its row, its selector column, and its record rows."""
    header_rows, info_cols = {}, {}
    for _, text in _markers(ws):
        if text.startswith("Header_"):
            index = int(text.split("_", 1)[1].split("=")[0].strip())
            header_rows[index] = None
        elif text.startswith("Info_"):
            left, right = text.split("=", 1)
            info_cols[int(left.split("_", 1)[1])] = right.strip()

    for row, text in _markers(ws):
        if text.startswith("Header_"):
            header_rows[int(text.split("_", 1)[1].split("=")[0].strip())] = row

    out = []
    for index in sorted(header_rows):
        header = header_rows[index]
        selector = info_cols.get(index)
        if header is None or selector is None:
            continue
        out.append({"index": index, "header": header, "selector": selector})
    return out


def _fields(ws, header_row):
    """The declared labels of one block, with the sheet's own words above them."""
    fields = []
    for col in range(2, ws.max_column + 1):
        declared = ws.cell(row=header_row, column=col).value
        if declared in (None, ""):
            continue
        text = str(declared)
        if text.startswith("←"):
            continue
        fields.append({
            "col": col,
            "declared": text,
            "source": ws.cell(row=header_row - 1, column=col).value,
        })
    return fields


def _records(ws, block, fields, last_row):
    """Rows the selector marks, and the rows it does not — both are carried across."""
    from openpyxl.utils import column_index_from_string

    sel = column_index_from_string(block["selector"])
    taken, excluded = [], []
    for row in range(block["header"] + 1, last_row + 1):
        marker = ws.cell(row=row, column=sel).value
        cells = [ws.cell(row=row, column=f["col"]) for f in fields]
        if all(c.value in (None, "") for c in cells):
            if taken or excluded:
                break                       # the block has ended
            continue
        (taken if marker not in (None, "") else excluded).append(cells)
    return taken, excluded


def transpose_sheet(source, target_wb, title: str) -> bool:
    """Write ``source`` onto its side. ``False`` where there is nothing to flip."""
    blocks = _blocks(source)
    if not blocks:
        return False

    ws = target_wb.create_sheet(title)
    put(ws, "B1", f"{source['B1'].value or title} — transposed", title_f)
    put(ws, "B2", "The same figures on their side: labels down column C, records across. "
                  "Nothing was retyped, so any difference in the output is the tool's.",
        sub_f)

    structural, attributes = [], []
    for _, text in _markers(source):
        (structural if text.split("_")[0] in ("Header", "Info", "Transpose")
         else attributes).append(text)

    marker_rows, row = [], 4
    for text in attributes:
        marker_rows.append((row, text))
        row += 1

    row = max(row, 6)
    for block in blocks:
        index = block["index"]
        fields = _fields(source, block["header"])
        taken, excluded = _records(source, block, fields, source.max_row)
        if not fields:
            continue

        put(ws, f"{SOURCE_COL}{row}", f"— block {index} —", head_f)
        first_field_row = row + 1
        for offset, spec in enumerate(fields):
            r = first_field_row + offset
            put(ws, f"{SOURCE_COL}{r}", spec["source"], inert_f)
            put(ws, f"{LABEL_COL}{r}", spec["declared"], head_f, blue)
            for n, cells in enumerate(taken + excluded):
                cell = cells[offset]
                out = ws.cell(row=r, column=FIRST_RECORD_COL + n, value=cell.value)
                out.font = body_f
                out.number_format = cell.number_format
        last_field_row = first_field_row + len(fields) - 1

        info_row = last_field_row + 1
        put(ws, f"{LABEL_COL}{info_row}", "selector →", head_f)
        for n in range(len(taken)):
            cell = ws.cell(row=info_row, column=FIRST_RECORD_COL + n, value="=COLUMN()")
            cell.font, cell.fill = body_f, green

        marker_rows.append((info_row - 3, f"Transpose_{index}"))
        marker_rows.append((info_row - 2, f"Header_{index} = {LABEL_COL}"))
        marker_rows.append((info_row - 1, f"Info_{index} = {info_row}"))
        row = info_row + 3

    for r, text in marker_rows:
        put(ws, f"A{r}", text, marker_f, yellow)

    put(ws, f"{SOURCE_COL}{row + 1}",
        "Column B is the sheet's own wording and is never read; column C holds the "
        "declared labels. A record column without a selector is excluded and reused as "
        "an independent control.", sub_f)
    widths(ws, {"A": 34, "B": 22, "C": 20})
    for col in range(FIRST_RECORD_COL, FIRST_RECORD_COL + 60):
        ws.column_dimensions[col_letter(col)].width = 13
    return True


def main(argv=None) -> int:
    from datatransform.recalc import inject

    argv = list(argv or sys.argv[1:]) or ["Intake_FireCatFull_v1.xlsx"]
    source_path = ROOT / argv[0]
    out_path = ROOT / source_path.name.replace("_v1.xlsx", "_Transposed_v1.xlsx")

    src = load_workbook(source_path, data_only=False)
    values = load_workbook(source_path, data_only=True)
    out = Workbook()
    out.remove(out.active)

    # Sheet 00 is the frame and is copied across unchanged: what is being tested is
    # whether the *data* sheets survive the flip, not the nomenclature.
    frame = out.create_sheet(src.sheetnames[0])
    source00 = src[src.sheetnames[0]]
    for row in source00.iter_rows():
        for cell in row:
            if cell.value not in (None, ""):
                new = frame.cell(row=cell.row, column=cell.column, value=cell.value)
                new.font, new.number_format = body_f, cell.number_format
                if str(cell.value).startswith("⟦"):
                    new.font, new.fill = head_f, grey
    widths(frame, {"A": 3, "B": 26, "C": 46, "D": 18})

    flipped = []
    for title in src.sheetnames[1:]:
        if transpose_sheet(values[title], out, f"{title}_T"):
            flipped.append(title)

    out.save(out_path)
    result = inject(out_path)
    print(f"written: {out_path.name}")
    print(f"flipped: {len(flipped)} sheet(s) — {', '.join(flipped)}")
    print(f"cached : {result['injected']} formula value(s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
