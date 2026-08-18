#!/usr/bin/env python3
"""Generate the reference workbooks for specification_v1.md.

Five treaty shapes, from the same building blocks:

===================================== ==================================================
``Intake_v1.xlsx``                    Fire only — one per-risk section, both orientations
``Intake_Engineering_v1.xlsx``        Engineering — 03 and 04 on separate sheets
``Intake_EngineeringCombined_v1.xlsx``  the same treaty, both loss datasets in one list
``Intake_FireCat_v1.xlsx``            Fire + Cat — the two cat perils share one sheet
``Intake_FireEQWind_v1.xlsx``         Fire + Earthquake + Hurricane — a sheet each
===================================== ==================================================

Two pairs, each making the same point twice. The cat pair differs only in *where the
section-blocks sit*; the Engineering pair only in whether the two loss datasets share a
sheet. Layout is not meaning: both pairs are verified to produce identical results.

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
            "Loss basis", "Date format", "Occurrence year from", "As at"]

HEADERS_01 = ["Year", "Premium", "Incurred Losses"]
HEADERS_02 = ["Year", "EPI"]
HEADERS_03 = ["Year", "Name of Loss", "Loss amount", "Date of Loss"]
# An event may be reported once per underwriting year it touches, so Event ID — not the
# name and not the year — is what identifies the event — spec §9.2.1 S15.
HEADERS_04 = ["Year", "Event ID", "Event Name", "Loss amount", "Event Date",
              "Event End Date", "Number of Claims (optional)"]

TYPES = [
    ("Year", "text"), ("Premium", "number"), ("Incurred Losses", "number"),
    ("EPI", "number"), ("Band", "text"), ("Band from", "number"), ("Band to", "number"),
    ("Number of Risks", "number"), ("Exposure", "number"), ("Claim Reference", "text"),
    ("Name of Loss", "text"), ("Event ID", "text"), ("Event Name", "text"),
    ("Loss amount", "number"), ("Number of Claims", "number"),
    ("Date of Loss", "date"), ("Event Date", "date"), ("Event End Date", "date"),
]

VOCABULARY = [
    ("Year basis", "UW | Occurrence"),
    ("Premium basis", "GWP | GNPI | Written | Earned | Signed"),
    ("Loss basis", "Paid | Incurred"),
    ("PF transfer", "with clean cut | without clean cut | none"),
    ("Exposure basis", "Sum Insured | EML | PML | MPL"),
    ("Includes fac", "yes | no"),
    ("Layered business", "included | excluded"),
    ("Scale", "1 | 1,000 | 1,000,000"),
    ("Share basis", "100% | ceded only"),
    ("Date format", "ISO | DD.MM.YYYY | MM/DD/YYYY"),
    ("Occurrence year from", "Event Date | Event End Date"),
    ("Currency", "ISO 4217"),
]

PERIOD_ORDER = [(1, "est"), (2, "9 months"), (3, "re-est")]

MARKER_LEGEND = [
    ("Header_i", "row-wise: this row is block i's extraction row (declared labels only)"),
    ("Header_i = <col>", "transposed: this column is block i's extraction column"),
    ("Info_i = <col>", "row-wise: this column holds the record selectors, =ROW()"),
    ("Info_i = <row>", "transposed: this row holds the record selectors, =COLUMN()"),
    ("Transpose_i", "block i is transposed"),
    ("Dataset_i = <key>", "which dataset block i is — defaults to the sheet-name match"),
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
    ("R-10", "01.Year basis", "=", "04.Year basis", 0, "error", "cat",
     "history and cat losses must be on the same year basis"),
]

# Declared in the order a profile is read: what it is, where it sits, what it earns,
# then what it is made of. Step 2 writes the columns in exactly this order — spec §9.2 S2.
#
# All three band fields are optional, because a band arrives as two numeric columns, as
# one label ("1 - 10,000"), or as the columns with no label at all. Step 2 always ends up
# with both bounds: where they are absent it reads them off the label and says so, and
# where the label is absent the bounds name the band. What must not happen is neither.
HEADERS_05 = ["Band (optional)", "Band from (optional)", "Band to (optional)",
              "Premium", "Number of Risks", "Exposure"]

# A profile is a snapshot of a portfolio, and three things decide whether its figures are
# comparable with anything else: whose share it is, whether facultative business is in it,
# and whether layered risks were flattened. Declared, so the question gets answered.
ATTRS_05 = ["Section", "Currency", "Scale", "Share basis", "Exposure basis",
            "Includes fac", "Layered business", "As at"]

PROVISIONAL: list = []


# ══════════════════════════════════════════════════════════════ sheet 00
def build_sheet00(wb, *, treaty_type, sections, datasets, rules=RULES, full_key_rules=None):
    ws = wb.create_sheet("00. NC+Interdep", 0)

    put(ws, "B1", "00. Nomenclature & Interdependencies", title_f)
    put(ws, "B2", "The frame: what each sheet holds, what the names mean, what must tie. "
                  "Column A is intentionally empty — no markers in this sheet.", sub_f)
    put(ws, "B3", "Greyed rows are declared but not yet implemented end to end.", sub_f)

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
    for name, value, fmt in [("Actual year", 2025, "0"),
                             ("Treaty type", treaty_type, None),
                             ("Cedent", "Example Insurance SA", None),
                             ("Loss share warning", 0.20, "0%")]:
        put(ws, f"B{r}", name, body_f)
        put(ws, f"C{r}", value, body_f, yellow, fmt=fmt)
        r += 1
    put(ws, f"B{r + 1}", "Actual year is N: the expiring year; the renewal is N+1. "
                         "Sheet attributes override these; block attributes override "
                         "the sheet. Loss share warning is the point above which a "
                         "year's declared losses are flagged (§10.3).", sub_f)

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


def loss_sheet(wb, name, *, section, title, losses, extra_attrs=None):
    """03. Large Losses — one claim per row, on a per-risk section."""
    ws = wb.create_sheet(name)
    put(ws, "B1", title, title_f)

    entries = {
        2: f"Section = {section}",
        3: "Currency = USD",
        4: "Scale = 1,000",
        5: "Share basis = 100%",
        6: "H_Year basis = UW",
        7: "H_Loss basis = Incurred",
        8: "Date format = ISO",
        9: "Threshold = 500",
        20: "As at = 31.12.2025",
    }
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
        declared={"C": "Year", "D": "Name of Loss", "E": "Loss amount",
                  "F": "Date of Loss"},
        rows=rows,
        formats={"C": "@", "E": "#,##0", "F": "yyyy-mm-dd"},
        note_col="G",
    )
    widths(ws, {"A": 32, "B": 12, "C": 10, "D": 32, "E": 14, "F": 14, "G": 46, "J": 10})
    return ws


# 04. Cat Losses. One row is one event *in one underwriting year*: an event that runs
# over a renewal date hits two underwriting years and is reported twice, so the annual
# sum and the cost of the event are different questions — spec §9.2.1 S15.
CAT_SOURCE_LABELS = {"B": "Cat code", "C": "U/W Yr", "D": "Event",
                     "E": "Gross incurred", "F": "From", "G": "To", "H": "Claims"}
CAT_DECLARED = {"B": "Event ID", "C": "Year", "D": "Event Name", "E": "Loss amount",
                "F": "Event Date", "G": "Event End Date", "H": "Number of Claims"}
CAT_FORMATS = {"C": "@", "E": "#,##0", "F": "yyyy-mm-dd", "G": "yyyy-mm-dd",
               "H": "#,##0"}


def _cat_layout(with_claims: bool) -> dict:
    """The cat block's columns. ``Number of Claims`` sits in H and is optional — §8 F4."""
    def keep(columns):
        return columns if with_claims else {c: v for c, v in columns.items() if c != "H"}

    return dict(
        selector="J",
        source=keep(CAT_SOURCE_LABELS),
        declared=keep(CAT_DECLARED),
        formats=keep(CAT_FORMATS),
        note="I",
    )


def _event_rows(events):
    """Expand each event into one row per underwriting year it touches."""
    rows = []
    for event_id, name, start, end, split in events:
        for year, amount, claims in split:
            row = {"B": event_id, "C": year, "D": name, "E": amount,
                   "F": start, "G": end}
            if claims is not None:
                row["H"] = claims
            rows.append(row)
    return rows


def cat_sheet(wb, name, *, section, title, events, with_claims, extra_attrs=None):
    """04. Cat Losses — one sheet, one cat section."""
    ws = wb.create_sheet(name)
    put(ws, "B1", title, title_f)
    layout = _cat_layout(with_claims)

    entries = {
        2: f"Section = {section}",
        3: "Currency = USD",
        4: "Scale = 1,000",
        5: "Share basis = 100%",
        6: "H_Year basis = UW",
        7: "H_Loss basis = Incurred",
        8: "Date format = ISO",
        9: "Occurrence year from = Event Date",
        20: "As at = 31.12.2025",
    }
    entries.update(extra_attrs or {})
    entries[11] = "Header_1"
    entries[19] = f"Info_1 = {layout['selector']}"
    markers(ws, entries)

    record_block(
        ws,
        header_row=11,
        selector_col=layout["selector"],
        source_labels=layout["source"],
        declared=layout["declared"],
        rows=_event_rows(events),
        formats=layout["formats"],
        note_col=layout["note"],
    )
    widths(ws, {"A": 32, "B": 12, "C": 10, "D": 30, "E": 14, "F": 14, "G": 14,
                "H": 10, "I": 46, "J": 10})
    return ws


def combined_loss_sheet(wb, name, *, section, title, losses, events):
    """One loss list, two blocks — spec §6.5.

    An Engineering or Miscellaneous cedent often reports large losses and cat events in
    a single list. Rather than splitting the sheet, two extraction rows are declared
    over the same records and **the selector column does the classifying**:

    ``J: =IF($C{row}="", ROW(), "")``   a row with no cat code is a large loss
    ``K: =IF($C{row}<>"", ROW(), "")``  a row with one is a cat loss

    Nothing is inferred — the cedent's own cat-code column drives it, and a human wrote
    the formula that says so.
    """
    ws = wb.create_sheet(name)
    put(ws, "B1", title, title_f)

    markers(ws, {
        2: f"Section = {section}",
        3: "Currency = USD",
        4: "Scale = 1,000",
        5: "Share basis = 100%",
        6: "H_Year basis = UW",
        7: "H_Loss basis = Incurred",
        8: "Date format = ISO",
        9: "Threshold = 500",
        10: "Occurrence year from = Event Date",
        12: "Header_1",
        13: "Dataset_1 = 03 Large",
        14: "Header_2",
        15: "Dataset_2 = 04 Cat",
        26: "Info_1 = J",
        27: "Info_2 = K",
        28: "As at = 31.12.2025",
    })

    for ref, text in [("B11", "Claim no."), ("C11", "Cat code"), ("D11", "U/W Yr"),
                      ("E11", "Description"), ("F11", "Gross incurred"), ("G11", "From"),
                      ("H11", "To"), ("I11", "Claims")]:
        put(ws, ref, text, inert_f)
    put(ws, "K11", "← source header, never read", sub_f)

    # Two extraction rows over one list. Each names only the fields its dataset declares.
    for ref, text in [("D12", "Year"), ("E12", "Name of Loss"), ("F12", "Loss amount"),
                      ("G12", "Date of Loss")]:
        put(ws, ref, text, head_f, blue)
    put(ws, "K12", "← extraction row: large losses", sub_f)

    for ref, text in [("C14", "Event ID"), ("D14", "Year"), ("E14", "Event Name"),
                      ("F14", "Loss amount"), ("G14", "Event Date"),
                      ("H14", "Event End Date"), ("I14", "Number of Claims")]:
        put(ws, ref, text, head_f, green)
    put(ws, "K14", "← extraction row: cat events", sub_f)

    # One list, interleaved as the cedent sent it — ordered by year, not by kind.
    rows, n = [], 0
    for year, what, amount, when in losses:
        n += 1
        rows.append({"B": f"CL-{2200 + n}", "D": year, "E": what, "F": amount, "G": when})
    for event_id, event_name, start, end, split in events:
        for year, amount, claims in split:
            rows.append({"C": event_id, "D": year, "E": event_name, "F": amount,
                         "G": start, "H": end, "I": claims})
    rows.sort(key=lambda r: (str(r["D"]), r["G"]))

    first = 16
    for offset, row in enumerate(rows):
        r = first + offset
        for col, value in row.items():
            fmt = {"D": "@", "F": "#,##0", "G": "yyyy-mm-dd",
                   "H": "yyyy-mm-dd", "I": "#,##0"}.get(col)
            put(ws, f"{col}{r}", value, body_f, fmt=fmt)
        is_cat = "C" in row
        put(ws, f"J{r}", None if is_cat else f'=IF($C{r}="",ROW(),"")', mono_f)
        put(ws, f"K{r}", None if not is_cat else f'=IF($C{r}<>"",ROW(),"")', mono_f)
        # openpyxl writes no cached value; supply the one Excel would compute.
        if is_cat:
            put(ws, f"K{r}", f'=IF($C{r}<>"",ROW(),"")', mono_f)
        else:
            put(ws, f"J{r}", f'=IF($C{r}="",ROW(),"")', mono_f)

    put(ws, f"B{first + len(rows) + 1}",
        "One list. Column J selects the large losses, column K the cat events; both "
        "read the cedent's own cat-code column. No row is in both, and none is lost.",
        sub_f)

    widths(ws, {"A": 34, "B": 12, "C": 11, "D": 10, "E": 34, "F": 14, "G": 13,
                "H": 13, "I": 9, "J": 10, "K": 46})
    # The selector formulas the tool cannot evaluate — the builder knows the answers.
    cached = {}
    for offset, row in enumerate(rows):
        r = first + offset
        cached[(name, f"J{r}")] = "" if "C" in row else r
        cached[(name, f"K{r}")] = r if "C" in row else ""
    return ws, cached


def profile_sheet(wb, name, *, title, blocks_spec, with_bounds):
    """05. Risk Profiles — one block per line of business.

    ``with_bounds`` decides which of the two shapes the cedent sent: two numeric columns,
    or a single label that step 2 has to read the bounds off. Both occur; both are
    supported; only the second produces the "read off Band" note.
    """
    ws = wb.create_sheet(name)
    put(ws, "B1", title, title_f)

    source = {"B": "Band", "C": "From", "D": "To", "E": "Premium",
              "F": "Risks", "G": "Sum insured"}
    declared = {"B": "Band", "C": "Band from", "D": "Band to", "E": "Premium",
                "F": "Number of Risks", "G": "Exposure"}
    formats = {"C": "#,##0", "D": "#,##0", "E": "#,##0", "F": "#,##0", "G": "#,##0"}
    if not with_bounds:
        for column in ("C", "D"):
            source.pop(column)
            declared.pop(column)
            formats.pop(column)

    entries = {
        2: "Currency = USD",
        3: "Scale = 1,000",
        4: "Share basis = 100%",
        5: "Exposure basis = Sum Insured",
        6: "Includes fac = yes",
        7: "Layered business = excluded",
    }

    row = 11
    for index, spec in enumerate(blocks_spec, start=1):
        entries[row - 2] = f"Section_{index} = {spec['section']}"
        entries[row - 1] = f"As at_{index} = {spec['as_at']}"
        entries[row] = f"Header_{index}"
        put(ws, f"B{row - 3}", f"— {spec['section']} —", head_f)
        end = record_block(
            ws,
            header_row=row,
            selector_col="J",
            source_labels=source,
            declared=declared,
            rows=spec["rows"],
            formats=formats,
            note_col="H",
            total_cols=["E", "F", "G"],
        )
        entries[end + 1] = f"Info_{index} = J"
        row = end + 6

    markers(ws, entries)
    put(ws, f"B{row - 3}",
        "One profile per line of business. The top band is open at the top, so its upper "
        "bound is absent rather than zero.", sub_f)
    widths(ws, {"A": 32, "B": 22, "C": 12, "D": 12, "E": 12, "F": 10, "G": 14,
                "H": 44, "J": 10})
    return ws


def multi_section_sheet(wb, name, *, title, kind, blocks_spec):
    """One sheet carrying several section-blocks, stacked — spec §2.3.

    ``blocks_spec`` is a list of dicts with ``section`` and ``rows``. This is the shape
    a Cat sheet arrives in when Earthquake and Windstorm share one tab; three separate
    sheets produce exactly the same section-blocks.
    """
    ws = wb.create_sheet(name)
    put(ws, "B1", title, title_f)

    layouts = {
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
    }

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
        entries[10] = "Occurrence year from = Event Date"

    row = 13
    for index, spec in enumerate(blocks_spec, start=1):
        # A cat block declares its own columns: one section may report claim counts
        # while another does not, and the optional field is simply absent.
        layout = _cat_layout(spec.get("with_claims", False)) if kind == "cat" \
            else layouts[kind]
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
    widths(ws, {"A": 32, "B": 18, "C": 14, "D": 30, "E": 14, "F": 14, "G": 46,
                "H": 46, "I": 46, "J": 10, "K": 10, "L": 10})
    if kind == "cat":
        widths(ws, {"B": 12, "C": 10, "G": 14, "H": 10})
    return ws


# ══════════════════════════════════════════════════════════════ the variants
YEARS = ["2021", "2022", "2023", "2024", "2025", "2025 9 months"]
PERIODS = ["2025 est", "2025 9 months", "2025 re-est", "2026"]


def _epi(re_est, nine_months, est, renewal):
    return [est, nine_months, re_est, renewal]


# An Engineering or Miscellaneous treaty is *per risk* and still carries cat events:
# a flood or a hailstorm hits a construction site like any other risk. So one per-risk
# section reports both loss datasets, and §10.3 holds their sum against 01.
ENGINEERING_EVENTS = [
    ("EV-301", "Flood, Saxony", date(2021, 7, 14), date(2021, 7, 18),
     [("2021", 1150, 24)]),
    ("EV-302", "Hailstorm, Po Valley", date(2024, 6, 22), date(2024, 6, 22),
     [("2023", 900, 18), ("2024", 2600, 41)]),
]


def build_engineering():
    """An Engineering treaty — one per-risk section carrying large *and* cat losses."""
    wb = Workbook()
    wb.remove(wb.active)
    sections = [("Engineering", "per risk", ["01", "02", "03", "04", "05"])]
    datasets = [
        ("01. History", "01 History", HEADERS_01, ATTRS_01),
        ("02. EPI Projections", "02 EPI", HEADERS_02, ATTRS_02),
        ("03. Large Losses", "03 Large", HEADERS_03, ATTRS_03),
        ("04. Cat Losses", "04 Cat", HEADERS_04, ATTRS_04),
        ("05. Risk Profiles", "05 Profile", HEADERS_05, ATTRS_05),
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
        title="03. Large Losses — Engineering", losses=ENGINEERING_LOSSES,
    )
    cat_sheet(
        wb, "04. Cat Losses", section="Engineering",
        title="04. Cat Losses — Engineering", events=ENGINEERING_EVENTS,
        with_claims=True,
    )
    profile_sheet(
        wb, "05. Risk Profiles", title="05. Risk Profiles — Engineering",
        with_bounds=False,
        blocks_spec=[{"section": "Engineering", "as_at": "30.09.2025",
                      "rows": _profile_rows(PROFILE_ENGINEERING, False)}],
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

# id, name, start, end, [(underwriting year, loss, claims)]
# One event, several underwriting years: risks written in 2022 and in 2023 were both on
# cover when the Aegean earthquake struck, so the cedent reports it twice. Bettina runs
# across a renewal date, which is the same problem in its sharpest form.
EQ_EVENTS = [
    ("EV-101", "Aegean earthquake M6.8", date(2023, 2, 6), date(2023, 2, 6),
     [("2022", 400, 12), ("2023", 6800, 188)]),
    ("EV-102", "Central Italy earthquake", date(2021, 11, 19), date(2021, 11, 20),
     [("2021", 950, 31)]),
]
WIND_EVENTS = [
    ("EV-201", "Windstorm Eunice", date(2022, 2, 18), date(2022, 2, 19),
     [("2021", 1400, None), ("2022", 4000, None)]),
    ("EV-202", "Windstorm Bettina", date(2024, 12, 30), date(2025, 1, 2),
     [("2024", 1000, None), ("2025", 6600, None)]),
    ("EV-203", "Windstorm Kyrill II", date(2024, 12, 9), date(2024, 12, 10),
     [("2024", 1800, None)]),
]
FIRE_LOSSES = [
    ("2021", "Warehouse fire, Lyon", 1850, date(2021, 3, 14)),
    ("2023", "Factory fire, Rotterdam", 3400, date(2023, 1, 27)),
    ("2024", "Chemical plant fire, Basel", 2600, date(2024, 5, 19)),
    ("2025", "Cold store collapse, Milan", 1420, date(2025, 2, 11)),
]

# One risk profile per line of business. Bands in the sheet's own scale (1,000), so
# "> 25,000" is a sum insured above 25 million. The top band is open at the top: its
# upper bound is absent, never zero.
#
# label, from, to, premium, risks, exposure
PROFILE_FIRE = [
    ("0 – 1 000", 0, 1_000, 2_150, 620, 310_000),
    ("1 001 – 5 000", 1_001, 5_000, 4_900, 480, 1_290_000),
    ("5 001 – 10 000", 5_001, 10_000, 5_300, 260, 1_880_000),
    ("10 001 – 25 000", 10_001, 25_000, 4_700, 118, 2_010_000),
    ("> 25 000", 25_000, None, 3_100, 27, 1_140_000),
]
PROFILE_ENGINEERING = [
    ("0 – 1 000", 0, 1_000, 1_150, 210, 96_000),
    ("1 001 – 5 000", 1_001, 5_000, 2_800, 165, 430_000),
    ("5 001 – 20 000", 5_001, 20_000, 3_450, 78, 780_000),
    ("> 20 000", 20_000, None, 2_900, 27, 940_000),
]
PROFILE_EQ = [
    ("0 – 2 500", 0, 2_500, 1_050, 380, 240_000),
    ("2 501 – 10 000", 2_501, 10_000, 2_300, 310, 690_000),
    ("10 001 – 50 000", 10_001, 50_000, 2_750, 190, 1_420_000),
    ("> 50 000", 50_000, None, 1_480, 80, 980_000),
]
PROFILE_WIND = [
    ("0 – 2 500", 0, 2_500, 1_320, 450, 290_000),
    ("2 501 – 10 000", 2_501, 10_000, 2_650, 370, 810_000),
    ("10 001 – 50 000", 10_001, 50_000, 3_180, 240, 1_680_000),
    ("> 50 000", 50_000, None, 1_910, 90, 1_120_000),
]


def _profile_rows(profile, with_bounds: bool):
    rows = []
    for label, low, high, premium, risks, exposure in profile:
        row = {"B": label, "E": premium, "F": risks, "G": exposure}
        if with_bounds:
            row["C"] = low
            if high is not None:
                row["D"] = high          # the open top band has no upper bound
        rows.append(row)
    return rows


CAT_SECTIONS = [
    ("Fire", "per risk", ["01", "02", "03", "05"]),
    ("Earthquake", "cat", ["01", "02", "04", "05"]),
    ("Windstorm", "cat", ["01", "02", "04", "05"]),
]


def _history_rows(spec):
    return [
        {"B": spec["years"][i], "C": spec["policies"][i],
         "D": spec["premium"][i], "G": spec["incurred"][i]}
        for i in range(len(spec["years"]))
    ]


def _epi_rows(values):
    return [{"B": PERIODS[i], "C": values[i]} for i in range(len(PERIODS))]


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
        ("05. Risk Profiles", "05 Profile", HEADERS_05, ATTRS_05),
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
            {"section": "Earthquake", "rows": _event_rows(EQ_EVENTS),
             "with_claims": True},
            {"section": "Windstorm", "rows": _event_rows(WIND_EVENTS)},
        ],
    )
    profile_sheet(
        wb, "05. Risk Profiles", title="05. Risk Profiles — Fire, EQ and Windstorm",
        with_bounds=True,
        blocks_spec=[
            {"section": "Fire", "as_at": "30.09.2025",
             "rows": _profile_rows(PROFILE_FIRE, True)},
            {"section": "Earthquake", "as_at": "30.09.2025",
             "rows": _profile_rows(PROFILE_EQ, True)},
            {"section": "Windstorm", "as_at": "30.09.2025",
             "rows": _profile_rows(PROFILE_WIND, True)},
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
        ("05. Risk Profiles Fire", "05 Profile Fire", HEADERS_05, ATTRS_05),
        ("05. Risk Profiles EQ", "05 Profile EQ", HEADERS_05, ATTRS_05),
        ("05. Risk Profiles Wind", "05 Profile Wind", HEADERS_05, ATTRS_05),
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
    cat_sheet(wb, "04. Cat Losses EQ", section="Earthquake",
              title="04. Cat Losses — Earthquake", events=EQ_EVENTS, with_claims=True)
    cat_sheet(wb, "04. Cat Losses Wind", section="Windstorm",
              title="04. Cat Losses — Windstorm", events=WIND_EVENTS,
              with_claims=False)

    for sheet, section, profile in [
        ("05. Risk Profiles Fire", "Fire", PROFILE_FIRE),
        ("05. Risk Profiles EQ", "Earthquake", PROFILE_EQ),
        ("05. Risk Profiles Wind", "Windstorm", PROFILE_WIND),
    ]:
        profile_sheet(
            wb, sheet, title=f"05. Risk Profiles — {section}", with_bounds=True,
            blocks_spec=[{"section": section, "as_at": "30.09.2025",
                          "rows": _profile_rows(profile, True)}],
        )
    return wb, "Intake_FireEQWind_v1.xlsx"


ENGINEERING_LOSSES = [
    ("2021", "Turbine erection collapse, Gdansk", 1420, date(2021, 4, 8)),
    ("2023", "Tunnel boring machine damage, Oslo", 2310, date(2023, 3, 22)),
    ("2023", "Crane failure, Rotterdam yard", 860, date(2023, 10, 4)),
    ("2024", "Generator fire, Valencia plant", 1975, date(2024, 7, 30)),
    ("2025", "Cable laying barge grounding, Kiel", 1180, date(2025, 5, 12)),
]


def build_engineering_combined():
    """The same Engineering treaty, with the losses arriving in **one list**.

    Identical figures to ``Intake_Engineering_v1.xlsx``; the only difference is that
    ``03`` and ``04`` share a sheet and are told apart by their selector columns.
    """
    wb = Workbook()
    wb.remove(wb.active)
    sections = [("Engineering", "per risk", ["01", "02", "03", "04", "05"])]
    datasets = [
        ("01. History", "01 History", HEADERS_01, ATTRS_01),
        ("02. EPI Projections", "02 EPI", HEADERS_02, ATTRS_02),
        ("03. Losses", "03 Large", HEADERS_03, ATTRS_03),
        ("04. Cat Losses", "04 Cat", HEADERS_04, ATTRS_04),
        ("05. Risk Profiles", "05 Profile", HEADERS_05, ATTRS_05),
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
    _, cached = combined_loss_sheet(
        wb, "03. Losses", section="Engineering",
        title="03. Losses — Engineering (large losses and cat events in one list)",
        losses=ENGINEERING_LOSSES, events=ENGINEERING_EVENTS,
    )
    profile_sheet(
        wb, "05. Risk Profiles", title="05. Risk Profiles — Engineering",
        with_bounds=False,
        blocks_spec=[{"section": "Engineering", "as_at": "30.09.2025",
                      "rows": _profile_rows(PROFILE_ENGINEERING, False)}],
    )
    return wb, "Intake_EngineeringCombined_v1.xlsx", cached


def main():
    from datatransform.recalc import inject

    builders = (build_engineering, build_fire_cat, build_fire_eq_wind,
                build_engineering_combined)
    for builder in builders:
        built = builder()
        wb, name = built[0], built[1]
        overrides = built[2] if len(built) > 2 else {}
        path = ROOT / name
        wb.save(path)
        result = inject(path, [(sheet, ref, value)
                               for (sheet, ref), value in overrides.items()])
        print(f"{name:34} {len(wb.sheetnames)} sheets, "
              f"{result['injected']} formula values cached")


if __name__ == "__main__":
    main()
