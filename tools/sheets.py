#!/usr/bin/env python3
"""Reusable sheet builders for the reference workbooks.

Every generated workbook is assembled from the same pieces, so the four treaty shapes
differ only in their ⟦SECTIONS⟧ block and which sheets they carry.
"""

from __future__ import annotations

from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col_letter

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

STRUCTURAL = ("Header_", "Info_", "Transpose_")


def put(ws, ref, value, font=body_f, fill=None, fmt=None):
    cell = ws[ref]
    cell.value = value
    cell.font = font
    if fill:
        cell.fill = fill
    if fmt:
        cell.number_format = fmt
    return cell


def block_header(ws, row, anchor, columns):
    """An ⟦ANCHOR⟧ header followed by a column-label row."""
    put(ws, f"B{row}", anchor, head_f, grey)
    for i, label in enumerate(columns):
        put(ws, f"{col_letter(2 + i)}{row + 1}", label, head_f)
    return row + 2


def markers(ws, entries: dict):
    """Write column A. Structural tokens are bold; attributes are plain."""
    for row, text in entries.items():
        cell = ws.cell(row=row, column=1, value=text)
        cell.fill = yellow
        cell.font = marker_f if text.startswith(STRUCTURAL) else attr_f
        cell.border = box


def record_block(ws, *, header_row, selector_col, source_labels, declared, rows,
                 formats, note_col="H", total_cols=None):
    """One row-wise block: an inert source header, an extraction row, then records.

    ``declared`` maps a column letter to the label declared in sheet 00; every other
    column of the source is deliberately left out of the extraction row.
    """
    for col, text in source_labels.items():
        put(ws, f"{col}{header_row - 1}", text, inert_f)
    put(ws, f"{note_col}{header_row - 1}", "← source header, never read", sub_f)

    for col, label in declared.items():
        put(ws, f"{col}{header_row}", label, head_f, blue).border = box
    put(ws, f"{note_col}{header_row}",
        "← extraction row: blanks mean 'not extracted'", sub_f)

    first = header_row + 1
    for i, values in enumerate(rows):
        row = first + i
        for col, value in values.items():
            put(ws, f"{col}{row}", value, body_f, fmt=formats.get(col))
        cell = ws[f"{selector_col}{row}"]
        cell.value = "=ROW()"
        cell.font, cell.fill, cell.border = body_f, green, box
    last = first + len(rows) - 1

    total_row = last + 1
    # ``total_cols`` names the columns the cedent would actually total. A band bound is
    # a number but not a quantity, so a profile's own total row leaves the bounds alone.
    numeric = (total_cols if total_cols is not None
               else [c for c, f in formats.items() if f == "#,##0"])
    key_col = next(iter(declared))
    put(ws, f"{key_col}{total_row}", "Total", head_f)
    for col in numeric:
        cell = ws[f"{col}{total_row}"]
        cell.value = f"=SUM({col}{first}:{col}{last})"
        cell.font, cell.number_format = head_f, "#,##0"
    put(ws, f"{note_col}{total_row}",
        "← selector empty: excluded from extraction, reused as control", sub_f)
    return total_row


def widths(ws, spec: dict):
    for col, width in spec.items():
        ws.column_dimensions[col].width = width


# ── the dataset inventory ─────────────────────────────────────────────────────
# The register above it says what *this pack* carries. This block says what the
# standard defines, so a reader can tell a sheet that is absent from a sheet that
# was never specified — spec §2.
#
#   role · sheet · what it holds · state
INVENTORY = [
    ("01", "History", "Premium and loss history per section", "implemented"),
    ("02", "EPI Projections", "Estimated premium income, N and N+1", "implemented"),
    ("03", "Large Losses", "Individual large claims, per-risk sections", "implemented"),
    ("04", "Cat Losses", "Catastrophe events, cat sections", "implemented"),
    ("05", "Risk Profiles", "Banded exposure — the per-risk rating basis", "implemented"),
    ("06", "EQ Aggs", "Earthquake sums insured per cat zone", "implemented"),
    ("07", "Wind Aggs", "Windstorm sums insured per cat zone", "implemented"),
    ("08", "Splits", "The cedent's book split — sums insured, one block per section",
     "implemented"),
    ("09", "Rate Development", "The rate change a submission claims — the UW's own sheet",
     "implemented"),
    ("10", "Triangles", "Loss development triangles", "outstanding"),
    ("11", "Exchange rates", "FX rates for conversion between currencies", "outstanding"),
    ("20", "Summary", "Collected step-2 blocks — written, never read", "outstanding"),
]


def inventory_block(ws, row, note=None):
    """Write ⟦INVENTORY⟧ — every dataset the standard defines, not just this pack's."""
    r = block_header(ws, row, "⟦INVENTORY⟧", ["Role", "Sheet", "Holds", "State"])
    for role, sheet, holds, state in INVENTORY:
        put(ws, f"B{r}", role, body_f)
        put(ws, f"C{r}", sheet, body_f)
        put(ws, f"D{r}", holds, body_f)
        put(ws, f"E{r}", state, body_f, grey if state != "implemented" else None)
        r += 1
    put(ws, f"B{r + 1}", note or (
        "Every dataset the tool knows, whether or not this pack carries it. The register "
        "above lists what is here; ⟦SECTIONS⟧ says which roles each section expects, so a "
        "missing sheet reads as structure rather than as a gap."), sub_f)
    return r + 2


# ═══════════════════════════ cat aggregates and split tables — spec §2.5, §2.6
#
# Shared by every pack carrying 06/07/08. What lives here is the *layout*; which axes a
# dataset declares and which ratios apply is the workbook's business, not this module's.

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
ATTRS_06 = ["Section", "Currency", "Scale", "Share basis", "Exposure basis",
            "Zone scheme", "Coinsurance", "Deductible", "Standard deductible",
            "Limit basis", "Multi-location", "Includes fac", "Period", "As at"]

HEADERS_08 = ["Category"] + [f"{c} (optional)" for c in COVER] + ["Total (optional)"]
ATTRS_08 = ["Section", "Currency", "Scale", "Share basis", "Exposure basis",
            "Includes fac", "As at"]

AGG_VERSIONS = [
    ("2025 9 months", "30.09.2025"),
    ("2026 at inception", "01.01.2026"),
    ("2026 at expiry", "31.12.2026"),
]

# through it — but the tool never assumes that, it only reports what it finds.
GROWTH = {"2025 9 months": 1.0, "2026 at inception": 1.062, "2026 at expiry": 1.128}

MIX = {                # ResB  ResC  ResBI  ComB  ComC  ComBI  IndB  IndC  IndBI
    "res": (0.42, 0.11, 0.02, 0.20, 0.06, 0.04, 0.10, 0.03, 0.02),
    "com": (0.20, 0.05, 0.01, 0.38, 0.11, 0.08, 0.12, 0.03, 0.02),
    "ind": (0.14, 0.04, 0.01, 0.18, 0.05, 0.04, 0.36, 0.11, 0.07),
    "mix": (0.28, 0.07, 0.01, 0.29, 0.09, 0.06, 0.14, 0.04, 0.02),
}

# Mexican earthquake exposure concentrates in the Valley of Mexico (13a/14a) and on the
# Pacific and Gulf coasts; 48 is the largest single accumulation in this book.

def check_catalogue(base, zones, name):
    """The fixture must cover the declared zoning exactly — no gaps, no strays."""
    listed = [zone for zone, *_ in base]
    if listed != zones:
        missing = [z for z in zones if z not in listed]
        extra = [z for z in listed if z not in zones]
        raise SystemExit(f"{name}: missing {missing}, unexpected {extra}, "
                         f"or out of catalogue order")


def _rows(base, version, detail=True):
    """``detail`` is the level this cedent reports at — spec §2.6.

    ``True`` the finished nine, ``"cover"`` the three cover columns, ``False`` a single
    figure per zone. The underlying portfolio is the same either way, which is what makes
    the three comparable.
    """
    factor = GROWTH[version]
    out = []
    for zone, size, mix in base:
        scaled = [round(size * share * factor) for share in MIX[mix]]
        row = {"B": zone}
        if detail == "cover":
            for i, cover in enumerate(COVER):
                row[chr(ord("C") + i)] = sum(
                    scaled[j] for j in range(len(scaled)) if j % 3 == i)
        elif detail:
            for i, value in enumerate(scaled):
                row[chr(ord("C") + i)] = value
            row["L"] = sum(scaled)                   # the cedent supplies the total
        else:
            row["C"] = sum(scaled)
        out.append(row)
    return out



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

    if detail == "cover":
        # The cedent splits by cover but not by occupancy — spec §2.6 case c).
        source = {"B": "Zone", "C": "Buildings", "D": "Contents", "E": "BI"}
        declared = {"B": "Zone", "C": "Building", "D": "Content", "E": "BI"}
        value_cols = list("CDE")
    elif detail:
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
    for index, (period, as_at) in enumerate(AGG_VERSIONS, start=1):
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




# ═════════════════════════════════════════════════ sheet 09 — spec §2.7
# The underwriter's own sheet: the rate change a submission claims. Two shapes, because
# an underwriter may hold the rate *levels* and not the movement between them.
HEADERS_09 = ["Year", "Rate (optional)", "Rate change (optional)"]
ATTRS_09 = ["Scope", "Source", "Rate change basis", "As at"]


def rate_sheet(wb, name, *, title, blocks_spec):
    """One block per scope. ``rows`` may carry a rate, a change, or both."""
    ws = wb.create_sheet(name)
    put(ws, "B1", title, title_f)
    put(ws, "B2", "Entered by the underwriter, not the cedent — so Source says where each "
                  "figure came from. Where only the rate level is held, step 2 works the "
                  "change out from one year to the next.", sub_f)

    entries, row = {}, 8
    for index, spec in enumerate(blocks_spec, start=1):
        entries[row - 5] = f"Scope_{index} = {spec['scope']}"
        entries[row - 4] = f"Source_{index} = {spec['source']}"
        entries[row - 3] = f"Rate change basis_{index} = {spec.get('basis', 'nominal')}"
        entries[row - 2] = f"As at_{index} = 30.09.2025"
        entries[row] = f"Header_{index}"
        put(ws, f"B{row - 6}", f"— {spec['scope']} —", head_f)

        declared = {"B": "Year"}
        source = {"B": "U/W year"}
        if spec.get("rates"):
            declared["C"], source["C"] = "Rate", "Rate %o"
        if spec.get("changes"):
            col = "D" if spec.get("rates") else "C"
            declared[col], source[col] = "Rate change", "Change vs prior"

        rows = []
        for i, year in enumerate(spec["years"]):
            r = {"B": year}
            if spec.get("rates"):
                r["C"] = spec["rates"][i]
            if spec.get("changes"):
                r["D" if spec.get("rates") else "C"] = spec["changes"][i]
            rows.append(r)

        end = record_block(
            ws, header_row=row, selector_col="G",
            source_labels=source, declared=declared, rows=rows,
            formats={c: "0.000" if declared.get(c) == "Rate" else "0.0%"
                     for c in "CD" if c in declared},
            note_col="H", total_cols=[],
        )
        entries[end + 1] = f"Info_{index} = G"
        row = end + 8

    markers(ws, entries)
    widths(ws, {"A": 30, "B": 12, "C": 14, "D": 14, "G": 10, "H": 46})
    return ws
