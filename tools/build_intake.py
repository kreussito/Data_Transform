#!/usr/bin/env python3
"""Generate the reference workbooks for specification_v1.md.

Four treaty shapes, from the same building blocks:

===================================== ==================================================
``Intake_v1.xlsx``                    Fire only — one per-risk section, both orientations
``Intake_Engineering_v1.xlsx``        Engineering — one per-risk section
``Intake_FireCat_v1.xlsx``            Fire + Cat — the two cat perils share one sheet
``Intake_FireEQWind_v1.xlsx``         Fire + Earthquake + Hurricane — a sheet each
===================================== ==================================================

The last two differ only in *where the section-blocks sit*. That is the point: a
section is a block, and whether the blocks live in one sheet or three is immaterial.

Run from the repository root:  python tools/build_intake.py
"""

from __future__ import annotations

import sys
from datetime import date
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
    mono_f,
    prov_f,
    put,
    record_block,
    sub_f,
    title_f,
    widths,
    yellow,
)

ROOT = Path(__file__).resolve().parents[1]

ATTRS_01 = ["Section", "Currency", "Scale", "Share basis", "Year basis",
            "Premium basis", "Loss basis", "PF transfer", "As at"]
ATTRS_02 = ["Section", "Currency", "Scale", "Share basis", "Year basis",
            "Premium basis", "As at"]
ATTRS_03 = ["Section", "Currency", "Scale", "Share basis", "Year basis",
            "Loss basis", "Threshold", "Date format", "As at"]
ATTRS_04 = ["Section", "Currency", "Scale", "Share basis", "Year basis",
            "Loss basis", "Date format", "As at"]

HEADERS_01 = ["Year", "Premium", "Incurred Losses"]
HEADERS_02 = ["Year", "EPI"]
HEADERS_03 = ["Year", "Name of Loss", "Loss amount", "Date of Loss"]
HEADERS_04 = ["Year", "Event Name", "Loss amount", "Event Date"]

TYPES = [
    ("Year", "text"), ("Premium", "number"), ("Incurred Losses", "number"),
    ("EPI", "number"), ("Band", "text"), ("Number of Risks", "number"),
    ("Sum Insured", "number"), ("Claim Reference", "text"),
    ("Name of Loss", "text"), ("Event Name", "text"), ("Loss amount", "number"),
    ("Date of Loss", "date"), ("Event Date", "date"),
]

VOCABULARY = [
    ("Year basis", "UW | Occurrence"),
    ("Premium basis", "GWP | GNPI | Written | Earned | Signed"),
    ("Loss basis", "Paid | Incurred"),
    ("PF transfer", "with clean cut | without clean cut | none"),
    ("Exposure basis", "Sum Insured | EML | PML | MPL"),
    ("Scale", "1 | 1,000 | 1,000,000"),
    ("Share basis", "100% | ceded only"),
    ("Date format", "ISO | DD.MM.YYYY | MM/DD/YYYY"),
    ("Currency", "ISO 4217"),
]

PERIOD_ORDER = [(1, "est"), (2, "9 months"), (3, "re-est")]

MARKER_LEGEND = [
    ("Header_i", "row-wise: this row is block i's extraction row (declared labels only)"),
    ("Header_i = <col>", "transposed: this column is block i's extraction column"),
    ("Info_i = <col>", "row-wise: this column holds the record selectors, =ROW()"),
    ("Info_i = <row>", "transposed: this row holds the record selectors, =COLUMN()"),
    ("Transpose_i", "block i is transposed"),
    ("Section_i = <name>", "which section of the treaty block i belongs to"),
    ("<Attribute> = <value>", "an attribute of the sheet (unsuffixed) or block i (suffixed _i)"),
    ("H_<Attribute> = <val>", "the same, declared as a hypothesis — raises a query"),
]

# Rules reference a *role* (01, 02, 03, 04) so that one declaration serves every
# section of a treaty — spec §2.3.
RULES = [
    ("R-05", "01.Premium@{N} 9 months", "=", "02.EPI@{N} 9 months", 1, "error", "all",
     "same nine months reported in two sheets"),
    ("R-06", "01.Premium@{N}", "=", "02.EPI@{N} re-est", 1, "error", "all",
     "a full-year N in History is an estimate, not an actual"),
    ("R-07", "02.EPI@{N} re-est", ">=", "02.EPI@{N} 9 months", 0, "error", "all",
     "premium accrues; a re-estimate cannot fall below what is booked"),
    ("R-01", "SUM(03.Loss amount@{Y})", "<=", "01.Incurred Losses@{Y}", 1, "error",
     "per risk", "large losses cannot exceed total incurred for the same year"),
    ("R-02", "SUM(04.Loss amount@{Y})", "<=", "01.Incurred Losses@{Y}", 1, "error",
     "cat", "cat losses cannot exceed total incurred for the same year"),
    ("R-09", "01.Year basis", "=", "03.Year basis", 0, "error", "per risk",
     "history and large losses must be on the same year basis"),
]

PROVISIONAL = [
    ("05. Risk Profiles", "05 Profile",
     ["Band", "Number of Risks", "Sum Insured", "Premium"],
     ["Section", "Currency", "Scale", "Exposure basis", "As at"]),
]


# ══════════════════════════════════════════════════════════════ sheet 00
def build_sheet00(wb, *, treaty_type, sections, datasets, rules=RULES, full_key_rules=None):
    ws = wb.create_sheet("00. NC+Interdep", 0)

    put(ws, "B1", "00. Nomenclature & Interdependencies", title_f)
    put(ws, "B2", "The frame: what each sheet holds, what the names mean, what must tie. "
                  "Column A is intentionally empty — no markers in this sheet.", sub_f)
    put(ws, "B3", "Greyed rows are provisional sketches, not yet specified.", sub_f)

    for ref, text in [("B4", "Sheet name"), ("C4", "Key"),
                      ("D4", "Headers  →"), ("N4", "Attributes  →")]:
        put(ws, ref, text, head_f, grey)

    row = 5
    for sheet_name, key, headers, attrs in datasets:
        put(ws, f"B{row}", sheet_name, body_f)
        put(ws, f"C{row}", key, body_f)
        for i, h in enumerate(headers):
            put(ws, f"{chr(68 + i)}{row}", h, body_f, blue)
        for i, a in enumerate(attrs):
            ws.cell(row=row, column=14 + i, value=a).fill = yellow
            ws.cell(row=row, column=14 + i).font = body_f
        row += 1
    for sheet_name, key, headers, attrs in PROVISIONAL:
        put(ws, f"B{row}", sheet_name, prov_f)
        put(ws, f"C{row}", key, prov_f)
        for i, h in enumerate(headers):
            put(ws, f"{chr(68 + i)}{row}", h, prov_f)
        for i, a in enumerate(attrs):
            ws.cell(row=row, column=14 + i, value=a).font = prov_f
        row += 1

    # ── ⟦SECTIONS⟧ — what this treaty is made of
    r = block_header(ws, row + 2, "⟦SECTIONS⟧", ["Section", "Kind", "Datasets"])
    for name, kind, roles in sections:
        put(ws, f"B{r}", name, body_f, yellow)
        put(ws, f"C{r}", kind, body_f, blue)
        put(ws, f"D{r}", ", ".join(roles), body_f)
        r += 1
    put(ws, f"B{r + 1}",
        "A section is a block. Per-risk sections carry large losses (03); cat sections "
        "carry event losses (04) instead. Rules are scoped by kind and evaluated once "
        "per applicable section, so an absent dataset is 'not applicable', not a gap.",
        sub_f)

    # ── ⟦GLOBAL⟧
    r = block_header(ws, r + 3, "⟦GLOBAL⟧", ["Attribute", "Value"])
    for name, value in [("Actual year", 2025), ("Treaty type", treaty_type),
                        ("Cedent", "Example Insurance SA")]:
        put(ws, f"B{r}", name, body_f)
        put(ws, f"C{r}", value, body_f, yellow, fmt="0" if name == "Actual year" else None)
        r += 1
    put(ws, f"B{r + 1}", "Actual year is N: the expiring year; the renewal is N+1. "
                         "Sheet attributes override these; block attributes override "
                         "the sheet.", sub_f)

    # ── ⟦TYPES⟧
    r = block_header(ws, r + 3, "⟦TYPES⟧", ["Field", "Type"])
    for name, kind in TYPES:
        put(ws, f"B{r}", name, body_f)
        put(ws, f"C{r}", kind, body_f, blue)
        r += 1
    put(ws, f"B{r + 1}", "A field name carries one type across the whole workbook.", sub_f)

    # ── ⟦VOCABULARY⟧
    r = block_header(ws, r + 3, "⟦VOCABULARY⟧", ["Attribute", "Permitted values"])
    for name, values in VOCABULARY:
        put(ws, f"B{r}", name, body_f)
        put(ws, f"C{r}", values, body_f)
        r += 1

    # ── ⟦PERIOD ORDER⟧
    r = block_header(ws, r + 2, "⟦PERIOD ORDER⟧", ["Rank", "Suffix"])
    for rank, suffix in PERIOD_ORDER:
        put(ws, f"B{r}", rank, body_f, fmt="0")
        put(ws, f"C{r}", suffix, body_f, blue)
        r += 1
    put(ws, f"B{r + 1}", "Within one year the suffixes sort in this order, not "
                         "alphabetically. A bare year sorts first; an unlisted suffix "
                         "sorts last.", sub_f)

    # ── ⟦RULES⟧
    r = block_header(ws, r + 3, "⟦RULES⟧",
                     ["ID", "Left", "Rel", "Right", "Tolerance", "Severity",
                      "Scope", "Note"])
    for rid, left, rel, right, tol, sev, scope, note in (full_key_rules or rules):
        put(ws, f"B{r}", rid, head_f)
        put(ws, f"C{r}", left, mono_f)
        put(ws, f"D{r}", rel, head_f)
        put(ws, f"E{r}", right, mono_f)
        put(ws, f"F{r}", tol, body_f)
        put(ws, f"G{r}", sev, body_f)
        put(ws, f"H{r}", scope, body_f, blue)
        put(ws, f"I{r}", note, sub_f)
        r += 1
    put(ws, f"B{r + 1}",
        "Forms: <key or role>.<field>@<record> · <key>.<attribute> · "
        "SUM(<key>.<field>@<record>).  {N} resolves from ⟦GLOBAL⟧; {Y} expands once "
        "per year. Scope selects the sections a rule runs over.", sub_f)

    # ── ⟦MARKERS⟧
    r = block_header(ws, r + 3, "⟦MARKERS⟧  —  written in column A of every other sheet",
                     ["Marker", "Meaning"])
    for marker, meaning in MARKER_LEGEND:
        put(ws, f"B{r}", marker, mono_f, yellow)
        put(ws, f"C{r}", meaning, body_f)
        r += 1
    put(ws, f"B{r + 1}",
        "No Header_i anywhere in column A  →  nothing is extracted from that sheet.",
        head_f)

    widths(ws, {"A": 3, "B": 26, "C": 42, "D": 20, "E": 34, "F": 10, "G": 9,
                "H": 10, "I": 44})
    for col in "JKLM":
        ws.column_dimensions[col].width = 16
    for col in "NOPQRSTUV":
        ws.column_dimensions[col].width = 14
    return ws


# ══════════════════════════════════════════════════════ data sheet builders
def history_sheet(wb, name, *, section, title, years, premium, incurred,
                  policies=None, extra_attrs=None):
    ws = wb.create_sheet(name)
    put(ws, "B1", title, title_f)
    policies = policies or [None] * len(years)

    entries = {
        2: f"Section = {section}",
        3: "Currency = USD",
        4: "Scale = 1,000",
        5: "Share basis = 100%",
        6: "H_Year basis = UW",
        7: "Premium basis = GNPI",
        8: "H_Loss basis = Incurred",
        9: "H_PF transfer = with clean cut",
        18: "As at = 31.12.2025",
    }
    entries.update(extra_attrs or {})
    entries[11] = "Header_1"
    entries[17] = "Info_1 = L"
    markers(ws, entries)

    rows = [
        {"B": years[i], "C": policies[i], "D": premium[i], "G": incurred[i]}
        for i in range(len(years))
    ]
    record_block(
        ws,
        header_row=11,
        selector_col="L",
        source_labels={"B": "U/W Yr", "C": "Policies", "D": "Prem.", "G": "Incurred"},
        declared={"B": "Year", "D": "Premium", "G": "Incurred Losses"},
        rows=rows,
        formats={"B": "@", "C": "#,##0", "D": "#,##0", "G": "#,##0"},
    )
    widths(ws, {"A": 32, "B": 16, "C": 12, "D": 12, "E": 10, "F": 10, "G": 14,
                "H": 46, "L": 10})
    return ws


def epi_sheet(wb, name, *, section, title, periods, epi, shares=None, extra_attrs=None):
    ws = wb.create_sheet(name)
    put(ws, "B1", title, title_f)
    shares = shares or [1.0] * len(periods)

    entries = {
        2: f"Section = {section}",
        3: "Currency = USD",
        4: "Scale = 1,000",
        5: "Share basis = 100%",
        6: "H_Year basis = UW",
        7: "Premium basis = GNPI",
        16: "As at = 31.12.2025",
    }
    entries.update(extra_attrs or {})
    entries[10] = "Header_1"
    entries[15] = "Info_1 = K"
    markers(ws, entries)

    rows = [{"B": periods[i], "C": epi[i], "D": shares[i]} for i in range(len(periods))]
    record_block(
        ws,
        header_row=10,
        selector_col="K",
        source_labels={"B": "Period", "C": "EPI (net)", "D": "Share %"},
        declared={"B": "Year", "C": "EPI"},
        rows=rows,
        formats={"B": "@", "C": "#,##0", "D": "0%"},
        note_col="F",
    )
    widths(ws, {"A": 32, "B": 18, "C": 14, "D": 10, "E": 8, "F": 46, "K": 10})
    return ws


def loss_sheet(wb, name, *, section, title, losses, kind="large", extra_attrs=None):
    """``kind`` is ``large`` (03, per risk) or ``cat`` (04)."""
    ws = wb.create_sheet(name)
    put(ws, "B1", title, title_f)
    is_large = kind == "large"

    entries = {
        2: f"Section = {section}",
        3: "Currency = USD",
        4: "Scale = 1,000",
        5: "Share basis = 100%",
        6: "H_Year basis = UW",
        7: "H_Loss basis = Incurred",
        8: "Date format = ISO",
        20: "As at = 31.12.2025",
    }
    if is_large:
        entries[9] = "Threshold = 500"
    entries.update(extra_attrs or {})
    entries[11] = "Header_1"
    entries[19] = "Info_1 = J"
    markers(ws, entries)

    rows = [
        {"B": f"CL-{2100 + i}", "C": year, "D": what, "E": amount, "F": when}
        for i, (year, what, amount, when) in enumerate(losses)
    ]
    record_block(
        ws,
        header_row=11,
        selector_col="J",
        source_labels={"B": "Claim no.", "C": "U/W Yr", "D": "Description",
                       "E": "Gross incurred", "F": "Occurred"},
        declared={"C": "Year",
                  "D": "Name of Loss" if is_large else "Event Name",
                  "E": "Loss amount",
                  "F": "Date of Loss" if is_large else "Event Date"},
        rows=rows,
        formats={"C": "@", "E": "#,##0", "F": "yyyy-mm-dd"},
        note_col="G",
    )
    widths(ws, {"A": 32, "B": 12, "C": 10, "D": 32, "E": 14, "F": 14, "G": 46, "J": 10})
    return ws


def multi_section_sheet(wb, name, *, title, kind, blocks_spec):
    """One sheet carrying several section-blocks, stacked — spec §2.3.

    ``blocks_spec`` is a list of dicts with ``section`` and ``rows``. This is the shape
    a Cat sheet arrives in when Earthquake and Windstorm share one tab; three separate
    sheets produce exactly the same section-blocks.
    """
    ws = wb.create_sheet(name)
    put(ws, "B1", title, title_f)

    layout = {
        "history": dict(
            selector="L",
            source={"B": "U/W Yr", "C": "Policies", "D": "Prem.", "G": "Incurred"},
            declared={"B": "Year", "D": "Premium", "G": "Incurred Losses"},
            formats={"B": "@", "C": "#,##0", "D": "#,##0", "G": "#,##0"},
            note="H",
        ),
        "epi": dict(
            selector="K",
            source={"B": "Period", "C": "EPI (net)"},
            declared={"B": "Year", "C": "EPI"},
            formats={"B": "@", "C": "#,##0"},
            note="F",
        ),
        "cat": dict(
            selector="J",
            source={"B": "Event no.", "C": "U/W Yr", "D": "Event",
                    "E": "Gross incurred", "F": "Occurred"},
            declared={"C": "Year", "D": "Event Name", "E": "Loss amount",
                      "F": "Event Date"},
            formats={"C": "@", "E": "#,##0", "F": "yyyy-mm-dd"},
            note="G",
        ),
    }[kind]

    entries = {
        2: "Currency = USD",
        3: "Scale = 1,000",
        4: "Share basis = 100%",
        5: "H_Year basis = UW",
        6: "As at = 31.12.2025",
    }
    if kind in ("history", "epi"):
        entries[7] = "Premium basis = GNPI"
    if kind in ("history", "cat"):
        entries[8] = "H_Loss basis = Incurred"
    if kind == "cat":
        entries[9] = "Date format = ISO"

    row = 12
    for index, spec in enumerate(blocks_spec, start=1):
        entries[row - 2] = f"Section_{index} = {spec['section']}"
        entries[row] = f"Header_{index}"
        put(ws, f"B{row - 3}", f"— {spec['section']} —", head_f)
        end = record_block(
            ws,
            header_row=row,
            selector_col=layout["selector"],
            source_labels=layout["source"],
            declared=layout["declared"],
            rows=spec["rows"],
            formats=layout["formats"],
            note_col=layout["note"],
        )
        entries[end + 1] = f"Info_{index} = {layout['selector']}"
        row = end + 5

    markers(ws, entries)
    widths(ws, {"A": 32, "B": 18, "C": 14, "D": 32, "E": 14, "F": 14, "G": 46,
                "H": 46, "J": 10, "K": 10, "L": 10})
    return ws


# ══════════════════════════════════════════════════════════════ the variants
YEARS = ["2021", "2022", "2023", "2024", "2025", "2025 9 months"]
PERIODS = ["2025 est", "2025 9 months", "2025 re-est", "2026"]


def _epi(re_est, nine_months, est, renewal):
    return [est, nine_months, re_est, renewal]


def build_engineering():
    """An Engineering treaty — one per-risk section, the same shape as Fire."""
    wb = Workbook()
    wb.remove(wb.active)
    sections = [("Engineering", "per risk", ["01", "02", "03"])]
    datasets = [
        ("01. History", "01 History", HEADERS_01, ATTRS_01),
        ("02. EPI Projections", "02 EPI", HEADERS_02, ATTRS_02),
        ("03. Large Losses", "03 Large", HEADERS_03, ATTRS_03),
    ]
    build_sheet00(wb, treaty_type="Engineering", sections=sections, datasets=datasets)

    history_sheet(
        wb, "01. History", section="Engineering",
        title="01. History — Engineering",
        years=YEARS,
        policies=[410, 435, 452, 468, 480, 372],
        premium=[8200, 8640, 9100, 9480, 9900, 7420],
        incurred=[5100, 6350, 4820, 7240, 5680, 3910],
    )
    epi_sheet(
        wb, "02. EPI Projections", section="Engineering",
        title="02. EPI Projections — Engineering",
        periods=PERIODS, epi=_epi(9900, 7420, 9450, 10600),
        shares=[1.0, 0.75, 1.0, 1.0],
    )
    loss_sheet(
        wb, "03. Large Losses", section="Engineering",
        title="03. Large Losses — Engineering",
        losses=[
            ("2021", "Turbine erection collapse, Gdansk", 1420, date(2021, 4, 8)),
            ("2023", "Tunnel boring machine damage, Oslo", 2310, date(2023, 3, 22)),
            ("2023", "Crane failure, Rotterdam yard", 860, date(2023, 10, 4)),
            ("2024", "Generator fire, Valencia plant", 1975, date(2024, 7, 30)),
            ("2025", "Cable laying barge grounding, Kiel", 1180, date(2025, 5, 12)),
        ],
    )
    return wb, "Intake_Engineering_v1.xlsx"


FIRE_HISTORY = dict(
    years=YEARS, policies=[1240, 1310, 1395, 1460, 1505, 1180],
    premium=[15900, 16740, 17520, 18390, 19200, 14400],
    incurred=[14930, 10030, 12660, 11700, 10250, 7100],
)
EQ_HISTORY = dict(
    years=YEARS, policies=[860, 890, 915, 940, 960, 745],
    premium=[6200, 6450, 6700, 6980, 7250, 5430],
    incurred=[1200, 480, 8900, 620, 540, 410],
)
WIND_HISTORY = dict(
    years=YEARS, policies=[1020, 1055, 1090, 1125, 1150, 890],
    premium=[7400, 7690, 7980, 8300, 8600, 6440],
    incurred=[2300, 6800, 1450, 3200, 9100, 5300],
)

EQ_EVENTS = [
    ("2023", "Aegean earthquake M6.8", 7200, date(2023, 2, 6)),
    ("2021", "Central Italy earthquake", 950, date(2021, 11, 19)),
]
WIND_EVENTS = [
    ("2022", "Windstorm Eunice", 5400, date(2022, 2, 18)),
    ("2025", "Windstorm Bettina", 7600, date(2025, 1, 31)),
    ("2024", "Windstorm Kyrill II", 2400, date(2024, 12, 9)),
]
FIRE_LOSSES = [
    ("2021", "Warehouse fire, Lyon", 1850, date(2021, 3, 14)),
    ("2023", "Factory fire, Rotterdam", 3400, date(2023, 1, 27)),
    ("2024", "Chemical plant fire, Basel", 2600, date(2024, 5, 19)),
    ("2025", "Cold store collapse, Milan", 1420, date(2025, 2, 11)),
]

CAT_SECTIONS = [
    ("Fire", "per risk", ["01", "02", "03"]),
    ("Earthquake", "cat", ["01", "02", "04"]),
    ("Windstorm", "cat", ["01", "02", "04"]),
]


def _history_rows(spec):
    return [
        {"B": spec["years"][i], "C": spec["policies"][i],
         "D": spec["premium"][i], "G": spec["incurred"][i]}
        for i in range(len(spec["years"]))
    ]


def _epi_rows(values):
    return [{"B": PERIODS[i], "C": values[i]} for i in range(len(PERIODS))]


def _event_rows(events):
    return [
        {"B": f"EV-{3100 + i}", "C": year, "D": what, "E": amount, "F": when}
        for i, (year, what, amount, when) in enumerate(events)
    ]


def build_fire_cat():
    """Fire + Cat — the two cat perils share one sheet, as blocks."""
    wb = Workbook()
    wb.remove(wb.active)
    datasets = [
        ("01. History Fire", "01 History Fire", HEADERS_01, ATTRS_01),
        ("01. History Cat", "01 History Cat", HEADERS_01, ATTRS_01),
        ("02. EPI Fire", "02 EPI Fire", HEADERS_02, ATTRS_02),
        ("02. EPI Cat", "02 EPI Cat", HEADERS_02, ATTRS_02),
        ("03. Large Losses Fire", "03 Large Fire", HEADERS_03, ATTRS_03),
        ("04. Cat Losses", "04 Cat", HEADERS_04, ATTRS_04),
    ]
    build_sheet00(wb, treaty_type="Fire + Nat Cat", sections=CAT_SECTIONS,
                  datasets=datasets)

    history_sheet(wb, "01. History Fire", section="Fire",
                  title="01. History — Fire section", **FIRE_HISTORY)
    multi_section_sheet(
        wb, "01. History Cat", title="01. History — Cat sections (EQ + Wind)",
        kind="history",
        blocks_spec=[
            {"section": "Earthquake", "rows": _history_rows(EQ_HISTORY)},
            {"section": "Windstorm", "rows": _history_rows(WIND_HISTORY)},
        ],
    )
    epi_sheet(wb, "02. EPI Fire", section="Fire",
              title="02. EPI Projections — Fire section",
              periods=PERIODS, epi=_epi(19200, 14400, 18500, 20900))
    multi_section_sheet(
        wb, "02. EPI Cat", title="02. EPI Projections — Cat sections (EQ + Wind)",
        kind="epi",
        blocks_spec=[
            {"section": "Earthquake", "rows": _epi_rows(_epi(7250, 5430, 7000, 7900))},
            {"section": "Windstorm", "rows": _epi_rows(_epi(8600, 6440, 8300, 9250))},
        ],
    )
    loss_sheet(wb, "03. Large Losses Fire", section="Fire",
               title="03. Large Losses — Fire section", losses=FIRE_LOSSES)
    multi_section_sheet(
        wb, "04. Cat Losses", title="04. Cat Losses — EQ + Wind",
        kind="cat",
        blocks_spec=[
            {"section": "Earthquake", "rows": _event_rows(EQ_EVENTS)},
            {"section": "Windstorm", "rows": _event_rows(WIND_EVENTS)},
        ],
    )
    return wb, "Intake_FireCat_v1.xlsx"


def build_fire_eq_wind():
    """Fire + Earthquake + Hurricane — a sheet per section."""
    wb = Workbook()
    wb.remove(wb.active)
    datasets = [
        ("01. History Fire", "01 History Fire", HEADERS_01, ATTRS_01),
        ("01. History EQ", "01 History EQ", HEADERS_01, ATTRS_01),
        ("01. History Wind", "01 History Wind", HEADERS_01, ATTRS_01),
        ("02. EPI Fire", "02 EPI Fire", HEADERS_02, ATTRS_02),
        ("02. EPI EQ", "02 EPI EQ", HEADERS_02, ATTRS_02),
        ("02. EPI Wind", "02 EPI Wind", HEADERS_02, ATTRS_02),
        ("03. Large Losses Fire", "03 Large Fire", HEADERS_03, ATTRS_03),
        ("04. Cat Losses EQ", "04 Cat EQ", HEADERS_04, ATTRS_04),
        ("04. Cat Losses Wind", "04 Cat Wind", HEADERS_04, ATTRS_04),
    ]
    build_sheet00(wb, treaty_type="Fire + Nat Cat", sections=CAT_SECTIONS,
                  datasets=datasets)

    history_sheet(wb, "01. History Fire", section="Fire",
                  title="01. History — Fire section", **FIRE_HISTORY)
    history_sheet(wb, "01. History EQ", section="Earthquake",
                  title="01. History — Earthquake section", **EQ_HISTORY)
    history_sheet(wb, "01. History Wind", section="Windstorm",
                  title="01. History — Windstorm section", **WIND_HISTORY)

    epi_sheet(wb, "02. EPI Fire", section="Fire",
              title="02. EPI Projections — Fire section",
              periods=PERIODS, epi=_epi(19200, 14400, 18500, 20900))
    epi_sheet(wb, "02. EPI EQ", section="Earthquake",
              title="02. EPI Projections — Earthquake section",
              periods=PERIODS, epi=_epi(7250, 5430, 7000, 7900))
    epi_sheet(wb, "02. EPI Wind", section="Windstorm",
              title="02. EPI Projections — Windstorm section",
              periods=PERIODS, epi=_epi(8600, 6440, 8300, 9250))

    loss_sheet(wb, "03. Large Losses Fire", section="Fire",
               title="03. Large Losses — Fire section", losses=FIRE_LOSSES)
    loss_sheet(wb, "04. Cat Losses EQ", section="Earthquake",
               title="04. Cat Losses — Earthquake", losses=EQ_EVENTS, kind="cat")
    loss_sheet(wb, "04. Cat Losses Wind", section="Windstorm",
               title="04. Cat Losses — Windstorm", losses=WIND_EVENTS, kind="cat")
    return wb, "Intake_FireEQWind_v1.xlsx"


def main():
    from datatransform.recalc import inject

    for builder in (build_engineering, build_fire_cat, build_fire_eq_wind):
        wb, name = builder()
        path = ROOT / name
        wb.save(path)
        result = inject(path)
        print(f"{name:32} {len(wb.sheetnames)} sheets, "
              f"{result['injected']} formula values cached")


if __name__ == "__main__":
    main()
