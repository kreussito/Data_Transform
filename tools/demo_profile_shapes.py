#!/usr/bin/env python3
"""One risk profile, three presentations — specification_v1.md §2.4.

The same portfolio, banded the same way, written down three different ways:

===================================== ==============================================
``05. Profile two columns``           Lower and Upper as separate numeric columns
``05. Profile one column``            one label per band: ``1-1,000,000``
``05. Profile European``              the same labels in ``1.000.000`` grouping
``05. Profile bounds only``           the two columns and **no label at all**
===================================== ==============================================

The point is that all three produce **identical** step-2 tables. Presentation is not
meaning: where the source supplies the bounds they are extracted, where it does not step 2
reads them off the label and says so, and either way the profile that comes out is the same.

Writes to ``output/`` — run from the repository root:

    python tools/demo_profile_shapes.py
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
    sub_f,
    title_f,
    widths,
    yellow,
)

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output" / "Demo_ProfileShapes.xlsx"

# One portfolio. Sums insured in full USD, so the bands are the ones you named.
# label parts:            lower        upper       premium    risks   exposure
PROFILE = [
    (1, 1_000_000, 1_240_000, 3_150, 812_000_000),
    (1_000_001, 5_000_000, 2_180_000, 1_420, 3_470_000_000),
    (5_000_001, 25_000_000, 2_760_000, 480, 5_240_000_000),
    (25_000_000, None, 1_820_000, 62, 3_180_000_000),
]

TYPES = [
    ("Band", "text"), ("Band from", "number"), ("Band to", "number"),
    ("Premium", "number"), ("Number of Risks", "number"), ("Exposure", "number"),
]

VOCABULARY = [
    ("Share basis", "100% | ceded only"),
    ("Exposure basis", "Sum Insured | EML | PML | MPL"),
    ("Includes fac", "yes | no"),
    ("Layered business", "included | excluded"),
    ("Scale", "1 | 1,000 | 1,000,000"),
    ("Currency", "ISO 4217"),
]


def anglo(value: int) -> str:
    return f"{value:,}"


def european(value: int) -> str:
    return f"{value:,}".replace(",", ".")


def label(lower, upper, group) -> str:
    return f"> {group(lower)}" if upper is None else f"{group(lower)}-{group(upper)}"


# ══════════════════════════════════════════════════════════════ sheet 00
def build_sheet00(wb):
    ws = wb.create_sheet("00. NC+Interdep", 0)
    put(ws, "B1", "00. Nomenclature & Interdependencies", title_f)
    put(ws, "B2", "One dataset, declared once. All three profile sheets carry the 05 "
                  "prefix, so all three resolve to it.", sub_f)

    for ref, text in [("B4", "Sheet name"), ("C4", "Key"),
                      ("D4", "Headers  →"), ("N4", "Attributes  →")]:
        put(ws, ref, text, head_f, grey)

    put(ws, "B5", "05. Risk Profiles", body_f)
    put(ws, "C5", "05 Profile", body_f)
    headers = ["Band (optional)", "Band from (optional)", "Band to (optional)",
               "Premium", "Number of Risks", "Exposure"]
    for i, text in enumerate(headers):
        put(ws, f"{chr(68 + i)}5", text, body_f, blue)
    attrs = ["Section", "Currency", "Scale", "Share basis", "Exposure basis",
             "Includes fac", "Layered business", "As at"]
    for i, text in enumerate(attrs):
        cell = ws.cell(row=5, column=14 + i, value=text)
        cell.font, cell.fill = body_f, yellow

    r = block_header(ws, 8, "⟦SECTIONS⟧", ["Section", "Kind", "Datasets"])
    put(ws, f"B{r}", "Fire", body_f, yellow)
    put(ws, f"C{r}", "per risk", body_f, blue)
    put(ws, f"D{r}", "05", body_f)

    r = block_header(ws, r + 3, "⟦GLOBAL⟧", ["Attribute", "Value"])
    for name, value in [("Actual year", 2025), ("Treaty type", "Fire"),
                        ("Cedent", "Example Insurance SA")]:
        put(ws, f"B{r}", name, body_f)
        put(ws, f"C{r}", value, body_f, yellow, fmt="0" if name == "Actual year" else None)
        r += 1

    r = block_header(ws, r + 2, "⟦TYPES⟧", ["Field", "Type"])
    for name, kind in TYPES:
        put(ws, f"B{r}", name, body_f)
        put(ws, f"C{r}", kind, body_f, blue)
        r += 1

    r = block_header(ws, r + 2, "⟦VOCABULARY⟧", ["Attribute", "Permitted values"])
    for name, values in VOCABULARY:
        put(ws, f"B{r}", name, body_f)
        put(ws, f"C{r}", values, body_f)
        r += 1

    widths(ws, {"A": 3, "B": 26, "C": 40, "D": 22, "E": 22, "F": 22, "G": 14,
                "H": 18, "I": 14})
    for col in "NOPQRSTU":
        ws.column_dimensions[col].width = 15
    return ws


# ══════════════════════════════════════════════════════════ profile sheets
def profile_sheet(wb, name, *, title, note, with_bounds, group=anglo,
                  with_label=True):
    ws = wb.create_sheet(name)
    put(ws, "B1", title, title_f)
    put(ws, "B2", note, sub_f)

    markers(ws, {
        3: "Section = Fire",
        4: "Currency = USD",
        5: "Scale = 1",
        6: "Share basis = 100%",
        7: "Exposure basis = Sum Insured",
        8: "Includes fac = yes",
        9: "Layered business = excluded",
        11: "Header_1",
        18: "Info_1 = J",
        19: "As at = 30.09.2025",
    })

    source = ["Band", "Lower", "Upper", "Premium", "Risks", "Sum insured"]
    declared = ["Band", "Band from", "Band to", "Premium", "Number of Risks", "Exposure"]
    columns = "BCDEFG"
    if not with_bounds:
        source = [source[0]] + source[3:]
        declared = [declared[0]] + declared[3:]
        columns = "BEFG"
    if not with_label:
        # The label column still holds the cedent's text; it is simply not declared in
        # the extraction row, so the tool never reads it. The bounds carry the identity.
        declared = [""] + declared[1:]

    for col, text in zip(columns, source):
        put(ws, f"{col}10", text, inert_f)
    put(ws, "H10", "← source header, never read", sub_f)
    for col, text in zip(columns, declared):
        if not text:
            continue
        put(ws, f"{col}11", text, head_f, blue).border = box
    put(ws, "H11",
        "← extraction row: the bounds are declared" if with_bounds
        else "← extraction row: no bounds here, so step 2 reads them off the label",
        sub_f)

    for i, (lower, upper, premium, risks, exposure) in enumerate(PROFILE):
        row = 12 + i
        put(ws, f"B{row}", label(lower, upper, group), body_f)
        if with_bounds:
            put(ws, f"C{row}", lower, body_f, fmt="#,##0")
            if upper is not None:
                put(ws, f"D{row}", upper, body_f, fmt="#,##0")
        put(ws, f"E{row}", premium, body_f, fmt="#,##0")
        put(ws, f"F{row}", risks, body_f, fmt="#,##0")
        put(ws, f"G{row}", exposure, body_f, fmt="#,##0")
        cell = ws[f"J{row}"]
        cell.value = "=ROW()"
        cell.font, cell.fill, cell.border = body_f, green, box
    if with_bounds:
        put(ws, f"D{11 + len(PROFILE)}", "open", inert_f)

    total_row = 12 + len(PROFILE)
    put(ws, f"B{total_row}", "Total", head_f)
    for col in "EFG":
        cell = ws[f"{col}{total_row}"]
        cell.value = f"=SUM({col}12:{col}{total_row - 1})"
        cell.font, cell.number_format = head_f, "#,##0"
    put(ws, f"H{total_row}", "← selector empty: excluded, reused as control", sub_f)

    widths(ws, {"A": 30, "B": 26, "C": 14, "D": 14, "E": 14, "F": 10, "G": 18,
                "H": 56, "J": 10})
    return ws


def build(out: Path = OUT):
    """The workbook, and the cached formula values its selectors need."""
    from datatransform.recalc import inject

    wb = Workbook()
    wb.remove(wb.active)
    build_sheet00(wb)

    profile_sheet(
        wb, "05. Profile two columns",
        title="05. Risk Profile — bounds in two columns",
        note="Lower = 1 and Upper = 1,000,000 sit in their own columns. Step 1 extracts "
             "them; step 2 has nothing to work out.",
        with_bounds=True,
    )
    profile_sheet(
        wb, "05. Profile one column",
        title="05. Risk Profile — bounds inside one label",
        note="The same bands written as '1-1,000,000'. Step 1 shows the label as it "
             "stands; step 2 reads the two bounds off it and says what it read.",
        with_bounds=False,
    )
    profile_sheet(
        wb, "05. Profile bounds only",
        title="05. Risk Profile — the two columns, no label declared",
        note="The Band column is left out of the extraction row entirely. Step 2 still "
             "produces both bounds, and names each band by them.",
        with_bounds=True, with_label=False,
    )
    profile_sheet(
        wb, "05. Profile European",
        title="05. Risk Profile — the same labels, European grouping",
        note="'1-1.000.000' rather than '1-1,000,000'. Both group thousands "
             "unambiguously, so both are read; nothing is guessed.",
        with_bounds=False, group=european,
    )

    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out, inject(out), wb.sheetnames


def main():
    out, result, sheets = build()
    print(f"written: {out}")
    print(f"sheets : {sheets}")
    print(f"cached : {result['injected']} formula value(s)")


if __name__ == "__main__":
    main()
