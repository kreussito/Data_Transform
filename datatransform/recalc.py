"""Write cached results next to formulas. Specification_v1.md §7.2.

``openpyxl`` emits formulas with an empty ``<v/>`` placeholder, so a workbook it wrote
reads back as ``None`` under ``data_only=True`` — which the spec treats as a hard stop,
because an uncached selector is indistinguishable from an empty cell.

Normally LibreOffice would recalculate. Where it is unavailable this fills the cache for
the formulas the tool itself emits: ``ROW()``, ``COLUMN()``, ``SUM(range)``, plus any
value supplied explicitly by the writer.
"""

from __future__ import annotations

import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ET.register_namespace("", NS)
Q = (lambda t: f"{{{NS}}}{t}")

REF = re.compile(r"^([A-Z]+)(\d+)$")
SHEET_XML = re.compile(r"xl/worksheets/sheet(\d+)\.xml$")


def col_to_num(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def split_ref(ref: str) -> tuple[int, int]:
    m = REF.match(ref)
    return col_to_num(m.group(1)), int(m.group(2))


def _expand(rng: str):
    a, b = rng.split(":")
    c1, r1 = split_ref(a)
    c2, r2 = split_ref(b)
    return [(c, r) for c in range(min(c1, c2), max(c1, c2) + 1)
            for r in range(min(r1, r2), max(r1, r2) + 1)]


def _process(xml_bytes: bytes, overrides: dict[str, object]):
    root = ET.fromstring(xml_bytes)

    literals = {}
    for c in root.iter(Q("c")):
        ref = c.get("r")
        if ref is None or c.find(Q("f")) is not None or c.get("t") in ("s", "str", "inlineStr"):
            continue
        v = c.find(Q("v"))
        if v is not None and v.text:
            try:
                literals[split_ref(ref)] = float(v.text)
            except ValueError:
                pass

    injected, unresolved = 0, []
    for c in root.iter(Q("c")):
        f = c.find(Q("f"))
        if f is None or not (f.text or "").strip():
            continue
        existing = c.find(Q("v"))
        if existing is not None and (existing.text or "").strip():
            continue

        ref = c.get("r")
        expr = f.text.strip().upper()
        val = overrides.get(ref)

        if val is None:
            col, row = split_ref(ref)
            if expr == "ROW()":
                val = row
            elif expr == "COLUMN()":
                val = col
            elif expr.startswith("SUM(") and expr.endswith(")") and ":" in expr:
                val = sum(literals.get(k, 0.0) for k in _expand(expr[4:-1]))
            else:
                unresolved.append(ref)
                continue

        v = existing if existing is not None else ET.SubElement(c, Q("v"))
        if isinstance(val, str):
            c.set("t", "str")
            v.text = val
        else:
            c.attrib.pop("t", None)
            v.text = repr(int(val)) if float(val).is_integer() else repr(float(val))
        injected += 1

    return ET.tostring(root, xml_declaration=True, encoding="UTF-8"), injected, unresolved


def inject(path: str | Path, formula_values=()) -> dict:
    """``formula_values`` is a sequence of ``(sheet_title, cell_ref, value)``."""
    path = Path(path)
    sheet_order = _sheet_order(path)
    by_sheet: dict[str, dict[str, object]] = {}
    for title, ref, value in formula_values:
        by_sheet.setdefault(title, {})[ref] = value

    backup = path.with_suffix(path.suffix + ".bak")
    shutil.copy2(path, backup)
    total, unresolved = 0, []
    try:
        with zipfile.ZipFile(backup) as zin, \
             zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                m = SHEET_XML.match(item.filename)
                if m:
                    title = sheet_order.get(item.filename, "")
                    data, n, missed = _process(data, by_sheet.get(title, {}))
                    total += n
                    unresolved.extend(f"{title}!{r}" for r in missed)
                zout.writestr(item, data)
    finally:
        backup.unlink(missing_ok=True)
    return {"injected": total, "unresolved": unresolved}


def _sheet_order(path: Path) -> dict[str, str]:
    """Map ``xl/worksheets/sheetN.xml`` to its sheet title via the workbook rels."""
    with zipfile.ZipFile(path) as z:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))

    rns = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    target = {r.get("Id"): r.get("Target") for r in rels.iter(f"{rns}Relationship")}
    dns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

    out = {}
    for sheet in wb.iter(Q("sheet")):
        tgt = target.get(sheet.get(f"{dns}id"), "").lstrip("/")
        if not tgt.startswith("xl/"):
            tgt = "xl/" + tgt
        out[tgt] = sheet.get("name")
    return out
