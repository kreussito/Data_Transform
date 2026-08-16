#!/usr/bin/env python3
"""Generate Intake_v1.xlsx — the reference workbook for specification_v1.md.

Run from the repository root:  python tools/build_intake.py
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Border, Font, PatternFill, Side

OUT = Path(__file__).resolve().parents[1] / "Intake_v1.xlsx"

FONT = "Arial"
title_f = Font(name=FONT, size=14, bold=True)
sub_f = Font(name=FONT, size=9, italic=True, color="666666")
head_f = Font(name=FONT, size=10, bold=True)
body_f = Font(name=FONT, size=10)
inert_f = Font(name=FONT, size=10, italic=True, color="999999")
marker_f = Font(name=FONT, size=10, bold=True, color="7F6000")
attr_f = Font(name=FONT, size=10, color="7F6000")
prov_f = Font(name=FONT, size=10, italic=True, color="999999")
mono_f = Font(name="Consolas", size=9)

yellow = PatternFill("solid", fgColor="FFF2CC")   # human-authored markers
green = PatternFill("solid", fgColor="E2EFDA")    # selector formulas
blue = PatternFill("solid", fgColor="DDEBF7")     # extraction row / column
grey = PatternFill("solid", fgColor="F2F2F2")     # block headers

thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()


def put(ws, ref, value, font=body_f, fill=None, fmt=None):
    cell = ws[ref]
    cell.value = value
    cell.font = font
    if fill:
        cell.fill = fill
    if fmt:
        cell.number_format = fmt
    return cell


def block(ws, row, anchor, columns):
    """An ⟦ANCHOR⟧ header followed by a column-label row."""
    put(ws, f"B{row}", anchor, head_f, grey)
    for i, label in enumerate(columns):
        put(ws, f"{chr(66 + i)}{row + 1}", label, head_f)
    return row + 2


# ══════════════════════════════════════════════════ 00. NC+Interdep
ws = wb.active
ws.title = "00. NC+Interdep"

put(ws, "B1", "00. Nomenclature & Interdependencies", title_f)
put(ws, "B2", "The frame: what each sheet holds, what the names mean, what must tie. "
              "Column A is intentionally empty — no markers in this sheet.", sub_f)
put(ws, "B3", "Rows 7-9 are provisional sketches, not yet specified.", sub_f)

# ── dataset register (rows 4-10, as agreed: dataset 01 on row 5)
for ref, txt in [("B4", "Sheet name"), ("C4", "Key"),
                 ("D4", "Headers  →"), ("N4", "Attributes  →")]:
    put(ws, ref, txt, head_f, grey)

ATTRS_01 = ["Currency", "Scale", "Year basis", "Premium basis",
            "Loss basis", "PF transfer", "As at"]
ATTRS_02 = ["Currency", "Scale", "Year basis", "Premium basis", "As at"]

datasets = [
    (5, "01. History", "01 History",
     ["Year", "Premium", "Incurred Losses"], ATTRS_01, False),
    (6, "02. EPI Projections", "02 EPI",
     ["Year", "EPI"], ATTRS_02, False),
    (7, "03. Large Losses", "03 Large",
     ["Year", "Loss Date", "Claim Reference", "Incurred Losses"],
     ["Currency", "Scale", "Year basis", "Loss basis", "Threshold", "As at"], True),
    (8, "04. Cat Losses", "04 Cat",
     ["Year", "Event Date", "Event Name", "Incurred Losses"],
     ["Currency", "Scale", "Year basis", "Loss basis", "As at"], True),
    (9, "05. Risk Profiles", "05 Profile",
     ["Band", "Number of Risks", "Sum Insured", "Premium"],
     ["Currency", "Scale", "Exposure basis", "As at"], True),
    (10, "01. History_Transposed", "01 History_T",
     ["Year", "Premium", "Incurred Losses"], ATTRS_01, False),
    (11, "02. EPI Projections_Transposed", "02 EPI_T",
     ["Year", "EPI"], ATTRS_02, False),
]

for row, sheet, key, headers, attrs, provisional in datasets:
    f = prov_f if provisional else body_f
    ws.cell(row=row, column=2, value=sheet).font = f
    ws.cell(row=row, column=3, value=key).font = f
    for i, h in enumerate(headers):                       # D..M
        c = ws.cell(row=row, column=4 + i, value=h)
        c.font = f
        if not provisional:
            c.fill = blue
    for i, a in enumerate(attrs):                         # N onward
        c = ws.cell(row=row, column=14 + i, value=a)
        c.font = f
        if not provisional:
            c.fill = yellow

# ── ⟦GLOBAL⟧
r = block(ws, 13, "⟦GLOBAL⟧", ["Attribute", "Value"])
for name, value in [("Actual year", 2025),
                    ("Cedent", "Example Insurance SA"),
                    ("Treaty", "Property per Risk XL")]:
    put(ws, f"B{r}", name, body_f)
    put(ws, f"C{r}", value, body_f, yellow, fmt="0" if name == "Actual year" else None)
    r += 1
put(ws, f"B{r + 1}", "Actual year is N: the expiring year. The renewal being underwritten "
                     "is N+1. Sheet attributes override these; block attributes override "
                     "the sheet.", sub_f)

# ── ⟦TYPES⟧
r = block(ws, r + 3, "⟦TYPES⟧", ["Field", "Type"])
types = [
    ("Year", "text"), ("Premium", "number"), ("Incurred Losses", "number"),
    ("EPI", "number"), ("Band", "text"), ("Number of Risks", "number"),
    ("Sum Insured", "number"), ("Claim Reference", "text"), ("Event Name", "text"),
    ("Loss Date", "text"), ("Event Date", "text"),
]
for name, kind in types:
    put(ws, f"B{r}", name, body_f)
    put(ws, f"C{r}", kind, body_f, blue)
    r += 1
put(ws, f"B{r + 1}", "A field name carries one type across the whole workbook. "
                     "Year is text because 2026 and '2026 9 months' are two records.", sub_f)

# ── ⟦VOCABULARY⟧
r = block(ws, r + 3, "⟦VOCABULARY⟧", ["Attribute", "Permitted values"])
for name, values in [
    ("Year basis", "UW | Occurrence"),
    ("Premium basis", "GWP | GNPI | Written | Earned | Signed"),
    ("Loss basis", "Paid | Incurred"),
    ("PF transfer", "with clean cut | without clean cut | none"),
    ("Exposure basis", "Sum Insured | EML | PML | MPL"),
    ("Scale", "1 | 1,000 | 1,000,000"),
    ("Currency", "ISO 4217"),
]:
    put(ws, f"B{r}", name, body_f)
    put(ws, f"C{r}", values, body_f)
    r += 1

# ── ⟦PERIOD ORDER⟧ — how suffixes sort within one leading number
r = block(ws, r + 2, "⟦PERIOD ORDER⟧", ["Rank", "Suffix"])
for rank, suffix in [(1, "est"), (2, "9 months"), (3, "re-est")]:
    put(ws, f"B{r}", rank, body_f, fmt="0")
    put(ws, f"C{r}", suffix, body_f, blue)
    r += 1
put(ws, f"B{r + 1}", "Within one year the suffixes sort in this order, not alphabetically: "
                     "est comes before 9 months even though '9' < 'e'. A bare year sorts "
                     "first; an unlisted suffix sorts last.", sub_f)

# ── ⟦RULES⟧
r = block(ws, r + 3, "⟦RULES⟧",
          ["ID", "Left", "Rel", "Right", "Tolerance", "Severity", "Note"])
rules = [
    ("R-05", "01 History.Premium@{N} 9 months", "=", "02 EPI.EPI@{N} 9 months",
     1, "error", "same nine months reported in two sheets"),
    ("R-06", "01 History.Premium@{N}", "=", "02 EPI.EPI@{N} re-est",
     1, "error", "a full-year N in History is an estimate, not an actual"),
    ("R-07", "02 EPI.EPI@{N} re-est", ">=", "02 EPI.EPI@{N} 9 months",
     0, "error", "premium accrues; a re-estimate cannot fall below what is booked"),
]
for rid, left, rel, right, tol, sev, note in rules:
    put(ws, f"B{r}", rid, head_f)
    put(ws, f"C{r}", left, mono_f)
    put(ws, f"D{r}", rel, head_f)
    put(ws, f"E{r}", right, mono_f)
    put(ws, f"F{r}", tol, body_f)
    put(ws, f"G{r}", sev, body_f)
    put(ws, f"H{r}", note, sub_f)
    r += 1
put(ws, f"B{r + 1}", "Form: <key>.<field>@<record>.  {N} resolves from ⟦GLOBAL⟧; records "
                     "match on leading number, then suffix. A rule is skipped, not "
                     "guessed, when Premium basis / Currency differ.", sub_f)

# ── ⟦MARKERS⟧
r = block(ws, r + 3, "⟦MARKERS⟧  —  written in column A of every other sheet",
          ["Marker", "Meaning"])
for m, d in [
    ("Header_i", "row-wise: this row is block i's extraction row (declared labels only)"),
    ("Header_i = <col>", "transposed: this column is block i's extraction column"),
    ("Info_i = <col>", "row-wise: this column holds the record selectors, =ROW()"),
    ("Info_i = <row>", "transposed: this row holds the record selectors, =COLUMN()"),
    ("Transpose_i", "block i is transposed"),
    ("<Attribute> = <value>", "an attribute of the sheet (unsuffixed) or block i (suffixed _i)"),
    ("H_<Attribute> = <val>", "the same, declared as a hypothesis — raises a query"),
]:
    put(ws, f"B{r}", m, marker_f, yellow)
    put(ws, f"C{r}", d, body_f)
    r += 1
put(ws, f"B{r + 1}", "No Header_i anywhere in column A  →  nothing is extracted from that sheet.",
    Font(name=FONT, size=10, bold=True))

ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 26
ws.column_dimensions["C"].width = 40
for col in "DEFGHIJKLM":
    ws.column_dimensions[col].width = 17
for col in "NOPQRST":
    ws.column_dimensions[col].width = 15


# ══════════════════════════════════════════════════ 01. History (row-wise)
ws = wb.create_sheet("01. History")

# 2025 appears twice: the full year, and the first nine months. The year alone does
# not identify a record — the full label does (spec §8.4, §9.2 S10).
years = [2021, 2022, 2023, 2024, 2025, "2025 9 months"]
policies = [1240, 1310, 1395, 1460, 1505, 1180]
premium = [15900, 16740, 17520, 18390, 19200, 14400]
comm = [0.225, 0.225, 0.230, 0.230, 0.235, 0.235]
rate = [1.02, 1.04, 1.06, 1.03, 1.05, 1.05]
incurred = [14930, 10030, 12660, 11700, 10250, 7100]
N1 = len(years)

put(ws, "B1", "01. History — row-wise reference block", title_f)

for row, txt in {
    2: "Currency = USD",
    3: "Scale = 1,000",
    4: "H_Year basis = UW",
    5: "Premium basis = GNPI",
    6: "H_Loss basis = Incurred",
    8: "Header_1",
    15: "H_PF transfer = with clean cut",
    17: "Info_1 = L",
    18: "As at = 31.12.2025",
}.items():
    c = ws.cell(row=row, column=1, value=txt)
    c.fill = yellow
    c.font = marker_f if txt.startswith(("Header_", "Info_", "Transpose_")) else attr_f
    c.border = box

put(ws, "B2", "Treaty XYZ — Property per Risk XL", Font(name=FONT, size=10, bold=True))

for col, txt in zip("BCDEFG", ["UW Year", "Policies", "Prem.", "Comm. %", "Rate", "Incurred"]):
    put(ws, f"{col}7", txt, inert_f)
put(ws, "H7", "← source header, never read", sub_f)
put(ws, "L7", "selector", sub_f)

for col, txt in [("B", "Year"), ("D", "Premium"), ("G", "Incurred Losses")]:
    put(ws, f"{col}8", txt, head_f, blue).border = box
put(ws, "H8", "← extraction row (Header_1): blanks mean 'not extracted'", sub_f)

for i in range(N1):
    r = 9 + i
    y = ws.cell(row=r, column=2, value=years[i])
    y.number_format = "@" if isinstance(years[i], str) else "0"
    for col, series, fmt in [(3, policies, "#,##0"), (4, premium, "#,##0"),
                             (5, comm, "0.0%"), (6, rate, "0.00"), (7, incurred, "#,##0")]:
        c = ws.cell(row=r, column=col, value=series[i])
        c.number_format = fmt
        c.font = body_f
    y.font = body_f
    s = ws.cell(row=r, column=12, value="=ROW()")
    s.font, s.fill, s.border = body_f, green, box

last = 8 + N1
put(ws, f"B{last + 1}", "Total", head_f)
for col, letter in [(3, "C"), (4, "D"), (7, "G")]:
    c = ws.cell(row=last + 1, column=col, value=f"=SUM({letter}9:{letter}{last})")
    c.font, c.number_format = head_f, "#,##0"
put(ws, f"H{last + 1}", "← selector empty: excluded from extraction, reused as control", sub_f)
put(ws, "B18", "Note: 2025 shown both as a full year and as the first nine months", sub_f)

ws.column_dimensions["A"].width = 34
ws.column_dimensions["B"].width = 15
for col in "CDEFG":
    ws.column_dimensions[col].width = 12
ws.column_dimensions["H"].width = 46
ws.column_dimensions["L"].width = 10


# ══════════════════════════════════════════════════ 01. History_Transposed
ws = wb.create_sheet("01. History_Transposed")

put(ws, "B1", "01. History — transposed reference block (same data, same result)", title_f)

for row, txt in {
    2: "Currency = USD",
    3: "Scale = 1,000",
    4: "H_Year basis = UW",
    5: "Transpose_1",
    6: "Header_1 = C",
    7: "Info_1 = 16",
    8: "Premium basis = GNPI",
    10: "H_Loss basis = Incurred",
    13: "H_PF transfer = with clean cut",
    16: "As at = 31.12.2025",
}.items():
    c = ws.cell(row=row, column=1, value=txt)
    c.fill = yellow
    c.font = marker_f if txt.startswith(("Header_", "Info_", "Transpose_")) else attr_f
    c.border = box

for row, txt in [(8, "UW Year"), (9, "Policies"), (10, "Prem."),
                 (11, "Comm. %"), (12, "Rate"), (13, "Incurred")]:
    ws.cell(row=row, column=2, value=txt).font = inert_f

for row, txt in [(8, "Year"), (10, "Premium"), (13, "Incurred Losses")]:
    c = ws.cell(row=row, column=3, value=txt)
    c.font, c.fill, c.border = head_f, blue, box

for row, series, fmt in [(8, years, "0"), (9, policies, "#,##0"), (10, premium, "#,##0"),
                         (11, comm, "0.0%"), (12, rate, "0.00"), (13, incurred, "#,##0")]:
    for i, v in enumerate(series):
        c = ws.cell(row=row, column=4 + i, value=v)
        c.number_format = "@" if isinstance(v, str) else fmt
        c.font = body_f

total_col = 4 + N1
last_letter = chr(64 + total_col - 1)
ws.cell(row=8, column=total_col, value="Total").font = head_f
for row in (9, 10, 13):
    c = ws.cell(row=row, column=total_col, value=f"=SUM(D{row}:{last_letter}{row})")
    c.font, c.number_format = head_f, "#,##0"

put(ws, "C16", "selector →", sub_f)
for i in range(N1):
    c = ws.cell(row=16, column=4 + i, value="=COLUMN()")
    c.font, c.fill, c.border = body_f, green, box

put(ws, "B18", "Column B is the sheet's own labels and is never read. "
               "Column C is the extraction column named by Header_1.", sub_f)
put(ws, "B19", "The Total column has no selector: excluded from extraction, "
               "reused as control.", sub_f)

ws.column_dimensions["A"].width = 34
ws.column_dimensions["B"].width = 14
ws.column_dimensions["C"].width = 18
for col in "DEFGHIJ":
    ws.column_dimensions[col].width = 14


# ══════════════════════════════════════════════════ 02. EPI Projections
ws = wb.create_sheet("02. EPI Projections")

# N = 2025, renewing 2026. The same year N carries three views, which is why the
# label and not the year identifies a record.
periods = ["2025 est", "2025 9 months", "2025 re-est", "2026"]
epi = [18500, 14400, 19200, 20900]
share = [1.00, 0.75, 1.00, 1.00]
comment = ["as advised at last renewal", "booked to 30.09", "current view", "renewal projection"]
N2 = len(periods)

put(ws, "B1", "02. EPI Projections — premium projections for N and N+1", title_f)

for row, txt in {
    2: "Currency = USD",
    3: "Scale = 1,000",
    4: "Premium basis = GNPI",
    5: "H_Year basis = UW",
    8: "Header_1",
    15: "Info_1 = K",
    16: "As at = 31.12.2025",
}.items():
    c = ws.cell(row=row, column=1, value=txt)
    c.fill = yellow
    c.font = marker_f if txt.startswith(("Header_", "Info_", "Transpose_")) else attr_f
    c.border = box

put(ws, "B6", "Treaty XYZ — Property per Risk XL", Font(name=FONT, size=10, bold=True))

# the sheet's own header, inert — different wording from the declared labels
for col, txt in zip("BCDE", ["Period", "EPI (net)", "Share %", "Comment"]):
    put(ws, f"{col}7", txt, inert_f)
put(ws, "F7", "← source header, never read", sub_f)
put(ws, "K7", "selector", sub_f)

# extraction row: EPI sits adjacent to Year here, unlike 01
for col, txt in [("B", "Year"), ("C", "EPI")]:
    put(ws, f"{col}8", txt, head_f, blue).border = box
put(ws, "F8", "← extraction row: Share % and Comment are not declared in 00", sub_f)

for i in range(N2):
    r = 9 + i
    put(ws, f"B{r}", periods[i], body_f, fmt="@")
    put(ws, f"C{r}", epi[i], body_f, fmt="#,##0")
    put(ws, f"D{r}", share[i], body_f, fmt="0%")
    put(ws, f"E{r}", comment[i], sub_f)
    s = ws.cell(row=r, column=11, value="=ROW()")       # column K
    s.font, s.fill, s.border = body_f, green, box

# an unmarked row: deliberately left to sheet 01, which carries the actuals
put(ws, "B13", "2024 actual", body_f, fmt="@")
put(ws, "C13", 17800, body_f, fmt="#,##0")
put(ws, "E13", "covered by 01. History", sub_f)
put(ws, "F13", "← selector empty: not extracted here", sub_f)

put(ws, "B17", "No total row: est, 9 months and re-est are three views of one year, "
               "so a sum of them would mean nothing.", sub_f)

ws.column_dimensions["A"].width = 30
ws.column_dimensions["B"].width = 18
ws.column_dimensions["C"].width = 14
ws.column_dimensions["D"].width = 10
ws.column_dimensions["E"].width = 30
ws.column_dimensions["F"].width = 44
ws.column_dimensions["K"].width = 10

# ══════════════════════════════════════════════════ 02. EPI Projections_Transposed
ws = wb.create_sheet("02. EPI Projections_Transposed")

put(ws, "B1", "02. EPI Projections — transposed (same data, same result)", title_f)

for row, txt in {
    2: "Currency = USD",
    3: "Scale = 1,000",
    4: "Premium basis = GNPI",
    5: "Transpose_1",
    6: "Header_1 = C",
    7: "Info_1 = 14",
    8: "H_Year basis = UW",
    10: "As at = 31.12.2025",
}.items():
    c = ws.cell(row=row, column=1, value=txt)
    c.fill = yellow
    c.font = marker_f if txt.startswith(("Header_", "Info_", "Transpose_")) else attr_f
    c.border = box

# column B — the sheet's own labels, inert
for row, txt in [(8, "Period"), (9, "Share %"), (10, "EPI (net)"), (11, "Comment")]:
    ws.cell(row=row, column=2, value=txt).font = inert_f

# column C — the extraction column
for row, txt in [(8, "Year"), (10, "EPI")]:
    c = ws.cell(row=row, column=3, value=txt)
    c.font, c.fill, c.border = head_f, blue, box

# columns D..G the records, H an unmarked one
labels = [*periods, "2024 actual"]
values = [*epi, 17800]
shares = [*share, 1.00]
notes = [*comment, "covered by 01. History"]
for i, label in enumerate(labels):
    col = 4 + i
    put(ws, f"{chr(64 + col)}8", label, body_f, fmt="@")
    put(ws, f"{chr(64 + col)}9", shares[i], body_f, fmt="0%")
    put(ws, f"{chr(64 + col)}10", values[i], body_f, fmt="#,##0")
    put(ws, f"{chr(64 + col)}11", notes[i], sub_f)

put(ws, "C14", "selector →", sub_f)
for i in range(len(periods)):                    # the 2024 actual column stays unmarked
    c = ws.cell(row=14, column=4 + i, value="=COLUMN()")
    c.font, c.fill, c.border = body_f, green, box

put(ws, "B16", "Column H (2024 actual) has no selector: not extracted here, "
               "because 01. History carries the actuals.", sub_f)

ws.column_dimensions["A"].width = 30
ws.column_dimensions["B"].width = 14
ws.column_dimensions["C"].width = 14
for col in "DEFGH":
    ws.column_dimensions[col].width = 17

wb.save(OUT)
print("written:", OUT)
print("sheets :", wb.sheetnames)
