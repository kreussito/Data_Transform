"""Sheet 00 — the nomenclature. Specification_v1.md §3."""

from __future__ import annotations

import re
import unicodedata

from .model import Dataset, ExtractionError

HEADER_COL_FIRST = 4     # D
HEADER_COL_LAST = 13     # M
ATTR_COL_FIRST = 14      # N
ATTR_COL_LAST = 40
SECTION_ROW = 4
FIRST_DATASET_ROW = 5

VOCAB_ANCHOR = "⟦VOCABULARY⟧"


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


class Nomenclature:
    """The dataset register and attribute vocabulary of sheet 00."""

    def __init__(self, datasets: dict[str, Dataset], vocabulary: dict[str, list[str]]):
        self.datasets = datasets
        self.vocabulary = vocabulary

    @classmethod
    def read(cls, ws) -> "Nomenclature":
        datasets: dict[str, Dataset] = {}
        for row in range(FIRST_DATASET_ROW, ws.max_row + 1):
            sheet_name = norm(ws.cell(row=row, column=2).value)
            key = norm(ws.cell(row=row, column=3).value)
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

        return cls(datasets, cls._read_vocabulary(ws))

    @staticmethod
    def _read_vocabulary(ws) -> dict[str, list[str]]:
        anchor = None
        for row in range(1, ws.max_row + 1):
            if norm(ws.cell(row=row, column=2).value) == VOCAB_ANCHOR:
                anchor = row
                break
        if anchor is None:
            return {}

        vocab: dict[str, list[str]] = {}
        for row in range(anchor + 2, ws.max_row + 1):     # skip the column-label row
            name = norm(ws.cell(row=row, column=2).value)
            values = norm(ws.cell(row=row, column=3).value)
            if not name:
                if vocab:
                    break
                continue
            if name.startswith("⟦"):
                break
            vocab[name] = [v.strip() for v in values.split("|") if v.strip()]
        return vocab

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
