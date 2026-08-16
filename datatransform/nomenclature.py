"""Sheet 00 — the nomenclature. Specification_v1.md §3."""

from __future__ import annotations

import re
import unicodedata

from .model import Dataset, ExtractionError, FieldType, Rule

HEADER_COL_FIRST = 4     # D
HEADER_COL_LAST = 13     # M
ATTR_COL_FIRST = 14      # N
ATTR_COL_LAST = 40
SECTION_ROW = 4
FIRST_DATASET_ROW = 5

GLOBAL_ANCHOR = "⟦GLOBAL⟧"
TYPES_ANCHOR = "⟦TYPES⟧"
VOCAB_ANCHOR = "⟦VOCABULARY⟧"
RULES_ANCHOR = "⟦RULES⟧"


def norm(value) -> str:
    """Trim, collapse whitespace, normalise NBSP and quote style — spec §8 F2, §2.2."""
    if value is None:
        return ""
    s = unicodedata.normalize("NFKC", str(value))
    s = s.replace(" ", " ").replace("„", '"').replace("“", '"').replace("”", '"')
    return re.sub(r"\s+", " ", s).strip()


def sheet_sort_key(name: str) -> tuple[str, str]:
    """Numeric prefix first, text second — spec §2.2."""
    n = norm(name)
    m = re.match(r"^(\d+)\s*[.\-]?\s*(.*)$", n)
    return (m.group(1).zfill(3), m.group(2).casefold()) if m else ("999", n.casefold())


def match_sheet(declared: str, available) -> str | None:
    """Numeric prefix first, text second — spec §2.2.

    The prefix is the stable identifier; the words after it are cosmetic, so
    ``10. Triangles`` resolves against a sheet actually spelled ``10. Triangels``.
    """
    exact = [s for s in available if norm(s).casefold() == norm(declared).casefold()]
    if exact:
        return exact[0]

    prefix, text = sheet_sort_key(declared)
    if prefix == "999":                       # no numeric prefix: nothing safe to fall back on
        return None
    candidates = [s for s in available if sheet_sort_key(s)[0] == prefix]
    if len(candidates) <= 1:
        return candidates[0] if candidates else None
    # Several sheets share the prefix — pick the closest text rather than guessing.
    return min(candidates, key=lambda s: _distance(sheet_sort_key(s)[1], text))


def _distance(a: str, b: str) -> int:
    """Levenshtein distance, used only to disambiguate a shared numeric prefix."""
    if a == b:
        return 0
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def _find_anchor(ws, anchor: str) -> int | None:
    for row in range(1, ws.max_row + 1):
        if norm(ws.cell(row=row, column=2).value) == anchor:
            return row
    return None


def _read_block(ws, anchor: str, columns: int, skip_label_row: bool = True):
    """Rows of an ⟦ANCHOR⟧ block, stopping at a blank row or the next anchor."""
    start = _find_anchor(ws, anchor)
    if start is None:
        return []

    rows = []
    for row in range(start + (2 if skip_label_row else 1), ws.max_row + 1):
        values = [norm(ws.cell(row=row, column=2 + i).value) for i in range(columns)]
        if not values[0]:
            break
        if values[0].startswith("⟦"):
            break
        rows.append((row, values))
    return rows


class Nomenclature:
    """Sheet 00: the frame — datasets, globals, types, vocabulary and rules."""

    def __init__(self, datasets, vocabulary, globals_=None, types=None, rules=None):
        self.datasets = datasets
        self.vocabulary = vocabulary
        self.globals = globals_ or {}
        self.types = types or {}
        self.rules = rules or []

    @property
    def actual_year(self) -> int | None:
        """N — the expiring year. The renewal being underwritten is N+1."""
        raw = self.globals.get("Actual year")
        if raw is None:
            return None
        try:
            return int(float(str(raw).replace(",", "")))
        except ValueError:
            raise ExtractionError(
                f"sheet 00 ⟦GLOBAL⟧: 'Actual year' is {raw!r}, which is not a year"
            ) from None

    def type_of(self, field: str) -> FieldType:
        declared = self.types.get(norm(field).casefold())
        if declared is None:
            raise ExtractionError(
                f"field {field!r} has no datatype in sheet 00 ⟦TYPES⟧ — "
                "every declared field needs one"
            )
        return declared

    def field_types(self, headers) -> dict[str, FieldType]:
        return {h: self.type_of(h) for h in headers}

    @classmethod
    def read(cls, ws) -> "Nomenclature":
        datasets: dict[str, Dataset] = {}
        for row in range(FIRST_DATASET_ROW, ws.max_row + 1):
            sheet_name = norm(ws.cell(row=row, column=2).value)
            key = norm(ws.cell(row=row, column=3).value)
            if sheet_name.startswith("⟦"):
                break                       # the register ends where the next block begins
            if not sheet_name or not key:
                continue
            headers = tuple(
                h for h in (
                    norm(ws.cell(row=row, column=c).value)
                    for c in range(HEADER_COL_FIRST, HEADER_COL_LAST + 1)
                ) if h
            )
            attributes = tuple(
                a for a in (
                    norm(ws.cell(row=row, column=c).value)
                    for c in range(ATTR_COL_FIRST, ATTR_COL_LAST + 1)
                ) if a
            )
            if not headers:
                continue
            datasets[sheet_name] = Dataset(sheet_name, key, headers, attributes)

        return cls(
            datasets,
            cls._read_vocabulary(ws),
            cls._read_globals(ws),
            cls._read_types(ws),
            cls._read_rules(ws),
        )

    @staticmethod
    def _read_globals(ws) -> dict[str, str]:
        return {name: value for _, (name, value) in _read_block(ws, GLOBAL_ANCHOR, 2)}

    @staticmethod
    def _read_types(ws) -> dict[str, FieldType]:
        return {
            name.casefold(): FieldType.parse(kind)
            for _, (name, kind) in _read_block(ws, TYPES_ANCHOR, 2)
        }

    @staticmethod
    def _read_rules(ws) -> list[Rule]:
        rules = []
        for row, values in _read_block(ws, RULES_ANCHOR, 7):
            rid, left, rel, right, tol, sev, note = values
            if not (left and rel and right):
                raise ExtractionError(
                    f"sheet 00 ⟦RULES⟧ row {row}: incomplete rule {rid!r}"
                )
            try:
                tolerance = float(str(tol).replace(",", "")) if tol else 0.0
            except ValueError:
                raise ExtractionError(
                    f"sheet 00 ⟦RULES⟧ row {row}: tolerance {tol!r} is not a number"
                ) from None
            rules.append(Rule(rid, left, rel, right, tolerance, sev or "error", note))
        return rules

    @staticmethod
    def _read_vocabulary(ws) -> dict[str, list[str]]:
        return {
            name: [v.strip() for v in values.split("|") if v.strip()]
            for _, (name, values) in _read_block(ws, VOCAB_ANCHOR, 2)
        }

    def dataset_for(self, sheet_name: str) -> Dataset | None:
        target = match_sheet(sheet_name, list(self.datasets))
        return self.datasets.get(target) if target else None

    def validate_value(self, attribute: str, value) -> bool:
        """An attribute with no vocabulary entry is unconstrained — spec §3.3."""
        allowed = self.vocabulary.get(attribute)
        if not allowed:
            return True
        if attribute == "Currency":
            return bool(re.fullmatch(r"[A-Z]{3}", norm(value).upper()))
        return norm(value).casefold() in {a.casefold() for a in allowed}


def read_nomenclature(wb) -> Nomenclature:
    name = match_sheet("00. Nomenclature & Interdependencies", wb.sheetnames)
    if name is None:
        name = next((s for s in wb.sheetnames if sheet_sort_key(s)[0] == "000"), None)
    if name is None:
        raise ExtractionError("sheet 00 not found — the nomenclature is required")
    return Nomenclature.read(wb[name])
