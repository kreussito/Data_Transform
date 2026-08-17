"""Column A marker grammar. Specification_v1.md §4.

Markers are *found, not positioned*: any row, any order. The ``_i`` suffix binds a
marker to block *i*; the ``H_`` prefix declares the value a hypothesis.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from .nomenclature import norm

# ``Section`` is deliberately not here: it is an attribute of the data (which part of
# the treaty these figures describe), and belongs in the attribute list a reviewer reads.
# ``Dataset`` is structure — which dataset the block *is* — so it stays out of that list.
STRUCTURAL = ("Header", "Info", "Transpose", "Dataset")

# Header_1 / Info_2 = L / Transpose_1 / Dataset_2 = 04 Cat / H_Year basis_1 = UW
MARKER = re.compile(
    r"^(?P<hyp>H_)?(?P<name>[^=]+?)(?:_(?P<index>\d+))?\s*(?:=\s*(?P<value>.*))?$"
)


@dataclass(frozen=True)
class Marker:
    name: str
    value: str | None
    index: int | None
    is_hypothesis: bool
    row: int

    @property
    def is_structural(self) -> bool:
        return self.name in STRUCTURAL


def parse_marker(text: str, row: int) -> Marker | None:
    s = norm(text)
    if not s:
        return None
    m = MARKER.match(s)
    if not m:
        return None

    name = m.group("name").strip()
    index = int(m.group("index")) if m.group("index") else None
    value = m.group("value")
    value = value.strip() if value is not None else None

    # A structural token never carries a hypothesis prefix.
    if name in STRUCTURAL and m.group("hyp"):
        name = "H_" + name

    return Marker(
        name=name,
        value=value,
        index=index,
        is_hypothesis=bool(m.group("hyp")) and name not in STRUCTURAL,
        row=row,
    )


def read_markers(ws, last_row: int | None = None) -> list[Marker]:
    """Every marker in column A of a data sheet."""
    limit = last_row or ws.max_row
    out = []
    for row in range(1, limit + 1):
        marker = parse_marker(ws.cell(row=row, column=1).value, row)
        if marker is not None:
            out.append(marker)
    return out


def block_indices(markers) -> list[int]:
    """Blocks are declared by ``Header_i`` and nothing else — spec §4 M7."""
    return sorted({m.index for m in markers if m.name == "Header" and m.index is not None})


def structural(markers, name: str, index: int) -> Marker | None:
    for m in markers:
        if m.name == name and m.index == index:
            return m
    return None


def attributes_for(markers, index: int) -> dict[str, Marker]:
    """Sheet-wide attributes, overridden by any carrying this block's suffix — spec §4 M5."""
    out: dict[str, Marker] = {}
    for m in markers:
        if m.is_structural or m.value is None:
            continue
        if m.index is None:
            out.setdefault(m.name, m)
    for m in markers:
        if m.is_structural or m.value is None:
            continue
        if m.index == index:
            out[m.name] = m
    return out
