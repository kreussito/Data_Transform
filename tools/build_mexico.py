#!/usr/bin/env python3
"""Generate Intake_Mexico_v1.xlsx — cat aggregates. Specification_v1.md §2.5.

A Mexican property treaty with two cat sections, each carrying a zone aggregate:

===================================== ==================================================
``06. EQ Aggs``                       Mexican earthquake zoning — 1 … 13a/13b, 14a…14d … 48
``07. Wind Aggs``                     Mexican hurricane zoning — 1 … 42
===================================== ==================================================

Each sheet carries the **three versions** a cedent sends — the nine-month estimate of the
expiring year, and projections for the first and last day of the renewal year — as three
stacked blocks sharing one selector column. What is interesting is not any single version
but the movement between them, and how that movement compares with premium growth.

Run from the repository root:  python tools/build_mexico.py
"""

from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.utils import get_column_letter as col_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from sheets import (  # noqa: E402
    block_header,
    blue,
    body_f,
    box,
    green,
    grey,
    head_f,
    inert_f,
    inventory_block,
    markers,
    put,
    record_block,
    sub_f,
    title_f,
    widths,
    yellow,
)

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "Intake_Mexico_v1.xlsx"

# ── the zone catalogues ───────────────────────────────────────────────────────
# Mexico's earthquake zoning replaces 13 with 13a/13b and 14 with 14a…14d.
EQ_ZONES = ([str(z) for z in range(1, 13)]
            + ["13a", "13b", "14a", "14b", "14c", "14d"]
            + [str(z) for z in range(15, 49)])
WIND_ZONES = [str(z) for z in range(1, 43)]

OCCUPANCY = ["Res", "Com", "Ind"]
COVER = ["Building", "Content", "BI"]

# The buckets are the product of the axes — spec §2.6. Earthquake declares both, so it
# has nine; windstorm declares occupancy alone, so it has three.
BUCKETS = [f"{o} {c}" for o in OCCUPANCY for c in COVER]

SEGMENTS = ["Projects", "Renewables"]

# Every level the aggregate may arrive at is declared, and all of them are optional: the
# register doubles as the list of shapes this dataset accepts. What is *not* optional is
# that one of them turns up — a block reporting none is refused by the split itself.
_LEVELS = [f"{o} (optional)" for o in OCCUPANCY] + \
          [f"{c} (optional)" for c in COVER] + \
          [f"{s} (optional)" for s in SEGMENTS] + ["Total (optional)"]
HEADERS_06 = ["Zone"] + [f"{b} (optional)" for b in BUCKETS] + _LEVELS
HEADERS_07 = ["Zone"] + _LEVELS

# Sheet 08 — the cedent's own split table, one per section (§2.6).
HEADERS_08 = ["Category"] + [f"{c} (optional)" for c in COVER] + ["Total (optional)"]
ATTRS_08 = ["Section", "Currency", "Scale", "Share basis", "Exposure basis",
            "Includes fac", "As at"]

AXES = [
    ("06 EQ Aggs", "Occupancy", OCCUPANCY),
    ("06 EQ Aggs", "Cover", COVER),
    ("07 Wind Aggs", "Occupancy", OCCUPANCY),
]

# Only the edges sheet 08 cannot supply. A cedent's split table gives every ratio
# between occupancy and cover; what no cedent cross-tabulates is Projects/Renewables
# against residential/commercial/industrial, and that a merged "Commercial" column
# covers industrial risks too.
SPLITS = [
    ("*", "Occupancy", "Renewables", "Ind", 1.00, "House convention"),
    ("*", "Occupancy", "Projects", "Com", 0.50, "House convention"),
    ("*", "Occupancy", "Projects", "Ind", 0.50, "House convention"),
    ("*", "Occupancy", "Commercial", "Com", 0.75, "House convention"),
    ("*", "Occupancy", "Commercial", "Ind", 0.25, "House convention"),
]

# ── sheet 08, per section ─────────────────────────────────────────────────────
# The book's sums insured on the grid the cedent keeps. Only ratios are read from it, so
# its scale and currency do not have to match 06/07 — only its own consistency matters.
#                     Category  Building  Content     BI
SPLIT_TABLE_EQ = [
    ("Res",             41_200,  10_800,    1_640),
    ("Com",             38_600,  11_900,    8_100),
    ("Ind",             22_400,   7_050,    4_900),
]
SPLIT_TABLE_WIND = [
    ("Res",             28_900,   7_600,    1_150),
    ("Com",             31_400,   9_700,    6_600),
    ("Ind",             15_100,   4_750,    3_300),
]

ATTRS_06 = ["Section", "Currency", "Scale", "Share basis", "Exposure basis",
            "Zone scheme", "Coinsurance", "Deductible", "Standard deductible",
            "Limit basis", "Multi-location", "Includes fac", "Period", "As at"]

HEADERS_01 = ["Year", "Premium", "Incurred Losses"]
ATTRS_01 = ["Section", "Currency", "Scale", "Share basis", "Year basis",
            "Premium basis", "Loss basis", "As at"]
HEADERS_02 = ["Year", "EPI"]
ATTRS_02 = ["Section", "Currency", "Scale", "Share basis", "Year basis",
            "Premium basis", "As at"]

TYPES = [("Year", "text"), ("Premium", "number"), ("Incurred Losses", "number"),
         ("EPI", "number"), ("Zone", "text"), ("Category", "text"),
         ("Total", "number")] + \
        [(b, "number") for b in BUCKETS] + [(o, "number") for o in OCCUPANCY] + \
        [(c, "number") for c in COVER] + [(s, "number") for s in SEGMENTS]

VOCABULARY = [
    ("Year basis", "UW | Occurrence"),
    ("Premium basis", "GWP | GNPI | Written | Earned | Signed"),
    ("Loss basis", "Paid | Incurred"),
    ("Share basis", "100% | ceded only"),
    ("Scale", "1 | 1,000 | 1,000,000"),
    ("Exposure basis", "Sum Insured | EML | PML | MPL"),
    ("Zone scheme", "Mexico EQ | Mexico Wind"),
    ("Coinsurance", "deducted | not deducted"),
    ("Deductible", "deducted | not deducted"),
    ("Limit basis", "full sum insured | per risk limit | per location limit"),
    ("Multi-location", "split by location | allocated to main zone | duplicated in each zone"),
    ("Includes fac", "yes | no"),
    ("Currency", "ISO 4217"),
]

# `at inception` and `at expiry` join the existing suffixes, so the three versions of an
# aggregate sort chronologically rather than alphabetically — spec §9.2 S13.
PERIOD_ORDER = [(1, "est"), (2, "9 months"), (3, "re-est"),
                (4, "at inception"), (5, "at expiry")]

SECTIONS = [
    ("Earthquake", "cat", ["01", "02", "06", "08"]),
    ("Hurricane", "cat", ["01", "02", "07", "08"]),
]

# Scale 1,000,000 — the figures are millions of MXN.
VERSIONS = [
    ("2025 9 months", "30.09.2025"),
    ("2026 at inception", "01.01.2026"),
    ("2026 at expiry", "31.12.2026"),
]

# ── the portfolio ─────────────────────────────────────────────────────────────
# The cedent returns the **complete** zone table, which is how a Mexican submission
# normally arrives: the zoning is the regulator's, so the form has a row per zone whether
# or not there is anything in it. A zone at 0 is therefore a statement — "we write nothing
# here" — and not a gap. Zones 7, 25 and 40 (EQ) and 14 and 36 (Wind) are those.
#
# A zone is described by its total sum insured and an occupancy pattern rather than by
# nine loose numbers, so the fixture reads as a portfolio and the nine buckets stay
# mutually consistent.
MIX = {                # ResB  ResC  ResBI  ComB  ComC  ComBI  IndB  IndC  IndBI
    "res": (0.42, 0.11, 0.02, 0.20, 0.06, 0.04, 0.10, 0.03, 0.02),
    "com": (0.20, 0.05, 0.01, 0.38, 0.11, 0.08, 0.12, 0.03, 0.02),
    "ind": (0.14, 0.04, 0.01, 0.18, 0.05, 0.04, 0.36, 0.11, 0.07),
    "mix": (0.28, 0.07, 0.01, 0.29, 0.09, 0.06, 0.14, 0.04, 0.02),
}

# Mexican earthquake exposure concentrates in the Valley of Mexico (13a/14a) and on the
# Pacific and Gulf coasts; 48 is the largest single accumulation in this book.
EQ_BASE = [
    ("1",  13_755, "mix"), ("2",   6_597, "res"), ("3",   2_450, "res"),
    ("4",   1_820, "res"), ("5",   3_960, "mix"), ("6",   5_240, "com"),
    ("7",       0, "res"), ("8",   2_180, "res"), ("9",   4_310, "mix"),
    ("10",  7_620, "com"), ("11",  1_540, "ind"), ("12",  3_070, "res"),
    ("13a", 25_519, "com"), ("13b", 9_049, "mix"),
    ("14a", 35_571, "com"), ("14b", 6_880, "mix"),
    ("14c",  5_393, "res"), ("14d", 2_940, "ind"),
    ("15",  4_120, "mix"), ("16",  1_260, "res"), ("17",  8_450, "ind"),
    ("18",  3_310, "com"), ("19",  2_070, "res"), ("20",  5_890, "mix"),
    ("21",  1_180, "res"), ("22", 18_671, "com"), ("23",  2_640, "ind"),
    ("24",  4_730, "mix"), ("25",      0, "res"), ("26",  3_180, "res"),
    ("27",  6_240, "com"), ("28",  1_920, "res"), ("29",  2_580, "mix"),
    ("30",  9_140, "ind"), ("31",  1_470, "res"), ("32",  3_860, "com"),
    ("33",  2_310, "mix"), ("34",  5_070, "res"), ("35",  2_796, "res"),
    ("36",  1_640, "ind"), ("37",  7_290, "com"), ("38",  2_150, "res"),
    ("39",  3_540, "mix"), ("40",      0, "res"), ("41",  4_680, "com"),
    ("42",  1_830, "res"), ("43",  2_960, "ind"), ("44",  6_410, "mix"),
    ("45",  1_390, "res"), ("46",  3_720, "com"), ("47",  2_240, "res"),
    ("48", 44_875, "com"),
]

# Hurricane exposure sits on the Caribbean and Gulf coasts — 3, 7, 19 and 42 carry it.
WIND_BASE = [
    ("1",   1_240, "res"), ("2",   3_570, "mix"), ("3",  16_014, "com"),
    ("4",   2_180, "res"), ("5",   5_320, "mix"), ("6",   1_760, "res"),
    ("7",  32_090, "com"), ("8",   4_610, "ind"), ("9",   2_240, "res"),
    ("10",  6_870, "mix"), ("11",  1_390, "res"), ("12",  7_878, "com"),
    ("13",  3_050, "mix"), ("14",      0, "res"), ("15",  2_640, "res"),
    ("16",  5_140, "com"), ("17",  1_820, "ind"), ("18",  3_960, "mix"),
    ("19", 23_208, "com"), ("20",  2_470, "res"), ("21",  6_310, "mix"),
    ("22",  1_580, "res"), ("23",  4_230, "ind"), ("24",  2_890, "com"),
    ("25",  1_340, "res"), ("26",  5_760, "mix"), ("27",  2_060, "res"),
    ("28",  4_052, "res"), ("29",  3_410, "com"), ("30",  1_670, "ind"),
    ("31",  6_540, "mix"), ("32",  2_320, "res"), ("33",  4_890, "com"),
    ("34",  1_450, "res"), ("35",  3_180, "mix"), ("36",      0, "res"),
    ("37",  5_430, "com"), ("38",  2_710, "res"), ("39",  1_960, "ind"),
    ("40",  4_360, "mix"), ("41",  2_150, "res"), ("42", 28_177, "com"),
]

# Version-to-version growth. The book grows into the renewal year and keeps growing
# through it — but the tool never assumes that, it only reports what it finds.
GROWTH = {"2025 9 months": 1.0, "2026 at inception": 1.062, "2026 at expiry": 1.128}


def _check_catalogue(base, zones, name):
    """The fixture must cover the declared zoning exactly — no gaps, no strays."""
    listed = [zone for zone, *_ in base]
    if listed != zones:
        missing = [z for z in zones if z not in listed]
        extra = [z for z in listed if z not in zones]
        raise SystemExit(f"{name}: missing {missing}, unexpected {extra}, "
                         f"or out of catalogue order")


def _rows(base, version, detail=True):
    """``detail`` distinguishes the two cedents: one sends the nine buckets, the other
    a single figure per zone and the book split separately."""
    factor = GROWTH[version]
    out = []
    for zone, size, mix in base:
        scaled = [round(size * share * factor) for share in MIX[mix]]
        row = {"B": zone}
        if detail:
            for i, value in enumerate(scaled):
                row[chr(ord("C") + i)] = value
            row["L"] = sum(scaled)                   # the cedent supplies the total
        else:
            row["C"] = sum(scaled)
        out.append(row)
    return out


# ══════════════════════════════════════════════════════════════════ sheet 00
def build_sheet00(wb):
    ws = wb.create_sheet("00. NC+Interdep", 0)
    put(ws, "B1", "00. Nomenclature & Interdependencies", title_f)
    put(ws, "B2", "The frame. Column A is intentionally empty — no markers in this sheet.",
        sub_f)
    put(ws, "B3", "The header register runs D…P here rather than D…M: 06 declares eleven "
                  "headers, and row 4 is what says where the boundary falls.", sub_f)

    datasets = [
        ("01. History EQ", "01 History EQ", HEADERS_01, ATTRS_01),
        ("01. History Wind", "01 History Wind", HEADERS_01, ATTRS_01),
        ("02. EPI EQ", "02 EPI EQ", HEADERS_02, ATTRS_02),
        ("02. EPI Wind", "02 EPI Wind", HEADERS_02, ATTRS_02),
        ("06. EQ Aggs", "06 EQ Aggs", HEADERS_06, ATTRS_06),
        ("07. Wind Aggs", "07 Wind Aggs", HEADERS_07, ATTRS_06),
        ("08. Splits EQ", "08 Splits EQ", HEADERS_08, ATTRS_08),
        ("08. Splits Wind", "08 Splits Wind", HEADERS_08, ATTRS_08),
    ]
    # Row 4 declares the boundary, and the boundary follows the widest register rather
    # than a column somebody once picked — §3.1. 06 alone declares nineteen headers.
    attr_col = 4 + max(len(headers) for _, _, headers, _ in datasets) + 1
    for ref, text in [("B4", "Sheet name"), ("C4", "Key"), ("D4", "Headers  →")]:
        put(ws, ref, text, head_f, grey)
    put(ws, f"{col_letter(attr_col)}4", "Attributes  →", head_f, grey)

    row = 5
    for sheet_name, key, headers, attrs in datasets:
        put(ws, f"B{row}", sheet_name, body_f)
        put(ws, f"C{row}", key, body_f)
        for i, h in enumerate(headers):
            cell = ws.cell(row=row, column=4 + i, value=h)
            cell.font, cell.fill = body_f, blue
        for i, a in enumerate(attrs):
            cell = ws.cell(row=row, column=attr_col + i, value=a)
            cell.font, cell.fill = body_f, yellow
        row += 1

    row = inventory_block(ws, row + 2)

    r = block_header(ws, row + 1, "⟦SECTIONS⟧", ["Section", "Kind", "Datasets"])
    for name, kind, roles in SECTIONS:
        put(ws, f"B{r}", name, body_f, yellow)
        put(ws, f"C{r}", kind, body_f, blue)
        put(ws, f"D{r}", ", ".join(roles), body_f)
        r += 1
    put(ws, f"B{r + 1}",
        "Either peril can be bought alone or both together, and the covered books need "
        "not be the same — so no rule compares 06 with 07.", sub_f)

    r = block_header(ws, r + 3, "⟦ZONES⟧", ["Scheme", "Zones"])
    for scheme, codes in [("Mexico EQ", EQ_ZONES), ("Mexico Wind", WIND_ZONES)]:
        put(ws, f"B{r}", scheme, body_f, yellow)
        put(ws, f"C{r}", ", ".join(codes), body_f)
        r += 1
    put(ws, f"B{r + 1}",
        "The complete zoning. This cedent returns every zone, so step 1 already shows all "
        "of them; where one lists only the zones it has exposure in, step 2 completes the "
        "rest with 0 — an absent zone reads as 'no data', a zero reads as 'nothing there'.",
        sub_f)

    r = block_header(ws, r + 3, "⟦GLOBAL⟧", ["Attribute", "Value"])
    for name, value, fmt in [("Actual year", 2025, "0"),
                             ("Treaty type", "Mexico property cat", None),
                             ("Cedent", "Aseguradora Ejemplo SA", None),
                             ("Loss share warning", 0.20, "0%"),
                             ("Rate change warning", 0.20, "0%")]:
        put(ws, f"B{r}", name, body_f)
        put(ws, f"C{r}", value, body_f, yellow, fmt=fmt)
        r += 1
    put(ws, f"B{r + 1}",
        "Rate change warning is the movement in premium per unit of exposure beyond "
        "which the growth block raises a warning (§10.4).", sub_f)

    r = block_header(ws, r + 3, "⟦TYPES⟧", ["Field", "Type"])
    for name, kind in TYPES:
        put(ws, f"B{r}", name, body_f)
        put(ws, f"C{r}", kind, body_f, blue)
        r += 1

    r = block_header(ws, r + 2, "⟦VOCABULARY⟧", ["Attribute", "Permitted values"])
    for name, values in VOCABULARY:
        put(ws, f"B{r}", name, body_f)
        put(ws, f"C{r}", values, body_f)
        r += 1

    r = block_header(ws, r + 2, "⟦PERIOD ORDER⟧", ["Rank", "Suffix"])
    for rank, suffix in PERIOD_ORDER:
        put(ws, f"B{r}", rank, body_f, fmt="0")
        put(ws, f"C{r}", suffix, body_f, blue)
        r += 1
    put(ws, f"B{r + 1}",
        "'at inception' and 'at expiry' order the three versions of an aggregate: the "
        "nine-month estimate of N, then the first and the last day of N+1.", sub_f)

    r = block_header(ws, r + 3, "⟦AXES⟧", ["Dataset", "Axis", "Categories"])
    for dataset, axis, categories in AXES:
        put(ws, f"B{r}", dataset, body_f)
        put(ws, f"C{r}", axis, body_f, blue)
        put(ws, f"D{r}", ", ".join(categories), body_f)
        r += 1
    put(ws, f"B{r + 1}",
        "The buckets of a dataset are the product of its axes: earthquake is occupancy "
        "× cover and has nine, hurricane declares occupancy alone and has three.", sub_f)

    r = block_header(ws, r + 3, "⟦SPLITS⟧",
                     ["Dataset", "Axis", "From", "Category", "Share", "Source"])
    for dataset, axis, frm, category, share, origin in SPLITS:
        put(ws, f"B{r}", dataset, body_f)
        put(ws, f"C{r}", axis, body_f, blue)
        put(ws, f"D{r}", frm, body_f)
        put(ws, f"E{r}", category, body_f)
        put(ws, f"F{r}", share, body_f, yellow, fmt="0%")
        put(ws, f"G{r}", origin, body_f)
        r += 1
    put(ws, f"B{r + 1}",
        "Ratios applied where the cedent reports at a coarser level. A blank 'From' "
        "distributes a whole axis; a filled one re-splits a category that arrived "
        "merged. 'Source' is not decoration — a ratio from the cedent's own prior "
        "submission and one borrowed from another book are both assumptions, but not "
        "equally good ones.", sub_f)

    widths(ws, {"A": 3, "B": 24, "C": 46, "D": 15})
    for col in range(5, attr_col):
        ws.column_dimensions[col_letter(col)].width = 14
    for col in range(attr_col, attr_col + 16):
        ws.column_dimensions[col_letter(col)].width = 15
    return ws


# ══════════════════════════════════════════════════════════════ aggregate sheet
def aggregate_sheet(wb, name, *, title, section, scheme, base, deductible, detail=True):
    """Three versions of one zone aggregate, stacked and sharing a selector column."""
    ws = wb.create_sheet(name)
    put(ws, "B1", title, title_f)
    put(ws, "B2", "Three versions of the same zoning. The movement between them is the "
                  "point — see the growth block written beneath.", sub_f)
    if not detail:
        put(ws, "B2", "Three versions, reported as one figure per zone. The occupancy "
                      "split comes from sheet 00 ⟦SPLITS⟧ — see the notes in step 2.",
            sub_f)

    entries = {
        3: f"Section = {section}",
        4: "Currency = MXN",
        5: "Scale = 1,000,000",
        6: "Share basis = 100%",
        7: "Exposure basis = Sum Insured",
        8: f"Zone scheme = {scheme}",
        9: "Coinsurance = not deducted",
        10: "Deductible = not deducted",
        11: f"Standard deductible = {deductible}",
        12: "Limit basis = full sum insured",
        13: "Multi-location = split by location",
        14: "Includes fac = yes",
    }

    if detail:
        source = {"B": "Zona", "C": "Res edif.", "D": "Res cont.", "E": "Res PB",
                  "F": "Com edif.", "G": "Com cont.", "H": "Com PB",
                  "I": "Ind edif.", "J": "Ind cont.", "K": "Ind PB", "L": "Suma"}
        declared = {"B": "Zone"}
        for i, bucket in enumerate(BUCKETS):
            declared[chr(ord("C") + i)] = bucket
        declared["L"] = "Total"
        value_cols = list("CDEFGHIJKL")
    else:
        source = {"B": "Zona", "C": "Suma asegurada"}
        declared = {"B": "Zone", "C": "Total"}
        value_cols = ["C"]
    formats = {c: "#,##0" for c in value_cols}

    row = 18
    for index, (period, as_at) in enumerate(VERSIONS, start=1):
        entries[row - 3] = f"Period_{index} = {period}"
        entries[row - 2] = f"As at_{index} = {as_at}"
        entries[row] = f"Header_{index}"
        put(ws, f"B{row - 4}", f"— {period} —", head_f)
        end = record_block(
            ws, header_row=row, selector_col="N",
            source_labels=source, declared=declared,
            rows=_rows(base, period, detail), formats=formats, note_col="O",
            total_cols=value_cols,
        )
        entries[end + 1] = f"Info_{index} = N"
        row = end + 7

    markers(ws, entries)
    widths(ws, {"A": 34, "B": 10, "N": 10, "O": 46})
    for col in value_cols:
        ws.column_dimensions[col].width = 14
    return ws


# ══════════════════════════════════════════════════════════════════ sheet 08
def split_sheet(wb, name, *, title, section, table):
    """The cedent's own split table — the book, not the zones. Spec §2.6."""
    ws = wb.create_sheet(name)
    put(ws, "B1", title, title_f)
    put(ws, "B2", "Sums insured for the whole book. 06/07 read only ratios from this, so "
                  "its scale and currency need not match theirs — only its own "
                  "consistency matters.", sub_f)
    markers(ws, {
        3: f"Section = {section}", 4: "Currency = MXN", 5: "Scale = 1,000,000",
        6: "Share basis = 100%", 7: "Exposure basis = Sum Insured",
        8: "Includes fac = yes", 9: "As at = 30.09.2025", 11: "Header_1", 16: "Info_1 = H",
    })
    rows = [{"B": category, "C": b, "D": c, "E": bi, "F": b + c + bi}
            for category, b, c, bi in table]
    record_block(
        ws, header_row=11, selector_col="H",
        source_labels={"B": "Ocupación", "C": "Edificio", "D": "Contenido",
                       "E": "Pérdida de beneficios", "F": "Suma"},
        declared={"B": "Category", "C": "Building", "D": "Content", "E": "BI",
                  "F": "Total"},
        rows=rows, formats={c: "#,##0" for c in "CDEF"}, note_col="I",
        total_cols=list("CDEF"),
    )
    widths(ws, {"A": 26, "B": 14, "H": 10, "I": 44})
    for col in "CDEF":
        ws.column_dimensions[col].width = 14
    return ws


# ══════════════════════════════════════════════════════════════ history sheets
def history_sheet(wb, name, *, title, section, premium, incurred):
    ws = wb.create_sheet(name)
    put(ws, "B1", title, title_f)
    markers(ws, {
        2: f"Section = {section}", 3: "Currency = MXN", 4: "Scale = 1,000,000",
        5: "Share basis = 100%", 6: "Year basis = UW", 7: "Premium basis = GNPI",
        8: "Loss basis = Incurred", 10: "Header_1", 19: "Info_1 = H",
        20: "As at = 30.09.2025",
    })
    years = ["2021", "2022", "2023", "2024", "2025", "2025 9 months"]
    record_block(
        ws, header_row=10, selector_col="H",
        source_labels={"B": "Año", "C": "Prima", "D": "Siniestros"},
        declared={"B": "Year", "C": "Premium", "D": "Incurred Losses"},
        rows=[{"B": years[i], "C": premium[i], "D": incurred[i]}
              for i in range(len(years))],
        formats={"B": "@", "C": "#,##0", "D": "#,##0"}, note_col="F",
        total_cols=["C", "D"],
    )
    widths(ws, {"A": 30, "B": 16, "C": 14, "D": 14, "F": 44, "H": 10})
    return ws


def epi_sheet(wb, name, *, title, section, epi):
    ws = wb.create_sheet(name)
    put(ws, "B1", title, title_f)
    markers(ws, {
        2: f"Section = {section}", 3: "Currency = MXN", 4: "Scale = 1,000,000",
        5: "Share basis = 100%", 6: "Year basis = UW", 7: "Premium basis = GNPI",
        9: "Header_1", 16: "Info_1 = G", 17: "As at = 30.09.2025",
    })
    periods = ["2025 est", "2025 9 months", "2025 re-est", "2026"]
    record_block(
        ws, header_row=9, selector_col="G",
        source_labels={"B": "Periodo", "C": "PPI"},
        declared={"B": "Year", "C": "EPI"},
        rows=[{"B": periods[i], "C": epi[i]} for i in range(len(periods))],
        formats={"B": "@", "C": "#,##0"}, note_col="E", total_cols=[],
    )
    widths(ws, {"A": 30, "B": 18, "C": 14, "E": 44, "G": 10})
    return ws


def main():
    from datatransform.recalc import inject

    _check_catalogue(EQ_BASE, EQ_ZONES, "EQ_BASE")
    _check_catalogue(WIND_BASE, WIND_ZONES, "WIND_BASE")

    wb = Workbook()
    wb.remove(wb.active)
    build_sheet00(wb)

    history_sheet(wb, "01. History EQ", title="01. History — Earthquake",
                  section="Earthquake",
                  premium=[1_180, 1_265, 1_340, 1_428, 1_512, 1_134],
                  incurred=[210, 185, 940, 240, 265, 198])
    history_sheet(wb, "01. History Wind", title="01. History — Hurricane",
                  section="Hurricane",
                  premium=[860, 918, 972, 1_035, 1_098, 823],
                  incurred=[340, 1_620, 295, 410, 380, 285])
    epi_sheet(wb, "02. EPI EQ", title="02. EPI Projections — Earthquake",
              section="Earthquake", epi=[1_480, 1_134, 1_512, 1_648])
    epi_sheet(wb, "02. EPI Wind", title="02. EPI Projections — Hurricane",
              section="Hurricane", epi=[1_070, 823, 1_098, 1_180])

    split_sheet(wb, "08. Splits EQ", title="08. Book split — Earthquake section",
                section="Earthquake", table=SPLIT_TABLE_EQ)
    split_sheet(wb, "08. Splits Wind", title="08. Book split — Hurricane section",
                section="Hurricane", table=SPLIT_TABLE_WIND)

    aggregate_sheet(wb, "06. EQ Aggs", title="06. Earthquake aggregates — Mexico",
                    section="Earthquake", scheme="Mexico EQ", base=EQ_BASE,
                    deductible="2%")
    aggregate_sheet(wb, "07. Wind Aggs", title="07. Hurricane aggregates — Mexico",
                    section="Hurricane", scheme="Mexico Wind", base=WIND_BASE,
                    deductible="1.5%", detail=False)

    wb.save(OUT)
    result = inject(OUT)
    print(f"written: {OUT}")
    print(f"sheets : {len(wb.sheetnames)} — {', '.join(wb.sheetnames)}")
    print(f"cached : {result['injected']} formula value(s)")


if __name__ == "__main__":
    main()
