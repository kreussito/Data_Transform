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

BUCKETS = ["Res Building", "Res Content", "Res BI",
           "Com Building", "Com Content", "Com BI",
           "Ind Building", "Ind Content", "Ind BI"]
HEADERS_06 = ["Zone"] + BUCKETS + ["Total (optional)"]

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
         ("EPI", "number"), ("Zone", "text"), ("Total", "number")] + \
        [(b, "number") for b in BUCKETS]

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
    ("Earthquake", "cat", ["01", "02", "06"]),
    ("Hurricane", "cat", ["01", "02", "07"]),
]

# Scale 1,000,000 — the figures are millions of MXN.
VERSIONS = [
    ("2025 9 months", "30.09.2025"),
    ("2026 at inception", "01.01.2026"),
    ("2026 at expiry", "31.12.2026"),
]

# Only the zones with exposure. Everything else is filled with 0 from ⟦ZONES⟧, which is
# the point: a zone with none is a fact worth seeing, not an absence.
#            zone   ResB   ResC  ResBI   ComB   ComC  ComBI   IndB   IndC  IndBI
EQ_BASE = [
    ("1",     4_820, 1_205,   180, 3_140,   940,   620, 1_880,   560,   410),
    ("2",     2_310,   578,    86, 1_505,   451,   298,   902,   270,   197),
    ("13a",   8_940, 2_235,   334, 5_820, 1_746, 1_152, 3_490, 1_040,   762),
    ("13b",   3_170,   792,   118, 2_064,   619,   409, 1_238,   369,   270),
    ("14a",  12_460, 3_115,   466, 8_112, 2_433, 1_606, 4_867, 1_450, 1_062),
    ("14c",   1_890,   472,    70, 1_230,   369,   243,   738,   220,   161),
    ("22",    6_540, 1_635,   244, 4_258, 1_277,   843, 2_555,   761,   558),
    ("35",     980,   245,    36,   638,   191,   126,   383,   114,    83),
    ("48",   15_720, 3_930,   588, 10_233, 3_069, 2_026, 6_140, 1_829, 1_340),
]
WIND_BASE = [
    ("3",     5_610, 1_402,   210, 3_652, 1_095,   723, 2_191,   653,   478),
    ("7",    11_240, 2_810,   421, 7_318, 2_195, 1_449, 4_391, 1_308,   958),
    ("12",    2_760,   690,   103, 1_797,   539,   355, 1_078,   321,   235),
    ("19",    8_130, 2_032,   304, 5_293, 1_587, 1_047, 3_176,   946,   693),
    ("28",    1_420,   355,    53,   924,   277,   183,   554,   165,   121),
    ("42",    9_870, 2_467,   370, 6_426, 1_927, 1_272, 3_856, 1_148,   841),
]

# Version-to-version growth. The book grows into the renewal year and keeps growing
# through it — but the tool never assumes that, it only reports what it finds.
GROWTH = {"2025 9 months": 1.0, "2026 at inception": 1.062, "2026 at expiry": 1.128}


def _rows(base, version):
    factor = GROWTH[version]
    out = []
    for zone, *buckets in base:
        scaled = [round(b * factor) for b in buckets]
        row = {"B": zone}
        for i, value in enumerate(scaled):
            row[chr(ord("C") + i)] = value
        row["L"] = sum(scaled)                       # the cedent supplies the total
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

    # Row 4 declares the boundary; the reader locates it rather than assuming — §3.1.
    for ref, text in [("B4", "Sheet name"), ("C4", "Key"),
                      ("D4", "Headers  →"), ("Q4", "Attributes  →")]:
        put(ws, ref, text, head_f, grey)

    datasets = [
        ("01. History EQ", "01 History EQ", HEADERS_01, ATTRS_01),
        ("01. History Wind", "01 History Wind", HEADERS_01, ATTRS_01),
        ("02. EPI EQ", "02 EPI EQ", HEADERS_02, ATTRS_02),
        ("02. EPI Wind", "02 EPI Wind", HEADERS_02, ATTRS_02),
        ("06. EQ Aggs", "06 EQ Aggs", HEADERS_06, ATTRS_06),
        ("07. Wind Aggs", "07 Wind Aggs", HEADERS_06, ATTRS_06),
    ]
    row = 5
    for sheet_name, key, headers, attrs in datasets:
        put(ws, f"B{row}", sheet_name, body_f)
        put(ws, f"C{row}", key, body_f)
        for i, h in enumerate(headers):
            cell = ws.cell(row=row, column=4 + i, value=h)
            cell.font, cell.fill = body_f, blue
        for i, a in enumerate(attrs):
            cell = ws.cell(row=row, column=17 + i, value=a)
            cell.font, cell.fill = body_f, yellow
        row += 1

    r = block_header(ws, row + 2, "⟦SECTIONS⟧", ["Section", "Kind", "Datasets"])
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
        "The complete zoning. A cedent lists only the zones it has exposure in; the rest "
        "are shown as 0, because an absent zone reads as 'no data' and a zero reads as "
        "'nothing there'.", sub_f)

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

    r = block_header(ws, r + 3, "⟦SPLITS⟧",
                     ["Scheme", "Occupancy", "Building", "Content", "BI"])
    put(ws, f"B{r + 1}",
        "Empty by design. Where a cedent sends fewer buckets than the nine, the ratios "
        "that fill the rest go here — declared, visible and versioned with the pack, "
        "because a split is an assumption and not a reading.", sub_f)

    widths(ws, {"A": 3, "B": 24, "C": 46, "D": 15})
    for col in "EFGHIJKLMNOP":
        ws.column_dimensions[col].width = 14
    for col in "QRSTUVWXYZ":
        ws.column_dimensions[col].width = 15
    return ws


# ══════════════════════════════════════════════════════════════ aggregate sheet
def aggregate_sheet(wb, name, *, title, section, scheme, base, deductible):
    """Three versions of one zone aggregate, stacked and sharing a selector column."""
    ws = wb.create_sheet(name)
    put(ws, "B1", title, title_f)
    put(ws, "B2", "Three versions of the same zoning. The movement between them is the "
                  "point — see the growth block written beneath.", sub_f)

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

    source = {"B": "Zona", "C": "Res edif.", "D": "Res cont.", "E": "Res PB",
              "F": "Com edif.", "G": "Com cont.", "H": "Com PB",
              "I": "Ind edif.", "J": "Ind cont.", "K": "Ind PB", "L": "Suma"}
    declared = {"B": "Zone"}
    for i, bucket in enumerate(BUCKETS):
        declared[chr(ord("C") + i)] = bucket
    declared["L"] = "Total"
    formats = {c: "#,##0" for c in "CDEFGHIJKL"}

    row = 18
    for index, (period, as_at) in enumerate(VERSIONS, start=1):
        entries[row - 3] = f"Period_{index} = {period}"
        entries[row - 2] = f"As at_{index} = {as_at}"
        entries[row] = f"Header_{index}"
        put(ws, f"B{row - 4}", f"— {period} —", head_f)
        end = record_block(
            ws, header_row=row, selector_col="N",
            source_labels=source, declared=declared,
            rows=_rows(base, period), formats=formats, note_col="O",
            total_cols=list("CDEFGHIJKL"),
        )
        entries[end + 1] = f"Info_{index} = N"
        row = end + 7

    markers(ws, entries)
    widths(ws, {"A": 34, "B": 10, "N": 10, "O": 46})
    for col in "CDEFGHIJKL":
        ws.column_dimensions[col].width = 12
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

    aggregate_sheet(wb, "06. EQ Aggs", title="06. Earthquake aggregates — Mexico",
                    section="Earthquake", scheme="Mexico EQ", base=EQ_BASE,
                    deductible="2%")
    aggregate_sheet(wb, "07. Wind Aggs", title="07. Hurricane aggregates — Mexico",
                    section="Hurricane", scheme="Mexico Wind", base=WIND_BASE,
                    deductible="1.5%")

    wb.save(OUT)
    result = inject(OUT)
    print(f"written: {OUT}")
    print(f"sheets : {len(wb.sheetnames)} — {', '.join(wb.sheetnames)}")
    print(f"cached : {result['injected']} formula value(s)")


if __name__ == "__main__":
    main()
