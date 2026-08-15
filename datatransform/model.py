"""Data model — specification_v1.md §4, §5, §6."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


class Orientation(str, Enum):
    ROW_WISE = "row-wise"
    TRANSPOSED = "transposed"


class Confidence(str, Enum):
    """Derived from the markers, never judged — spec §5.3.

    Ordered worst-last so that ``max`` yields the worst state of a set.
    """

    CONFIRMED = "Confirmed"
    ASSUMED = "Assumed"
    OPEN = "Open"

    @property
    def rank(self) -> int:
        return {"Confirmed": 0, "Assumed": 1, "Open": 2}[self.value]

    @staticmethod
    def worst(states) -> "Confidence":
        states = list(states)
        if not states:
            return Confidence.CONFIRMED
        return max(states, key=lambda s: s.rank)


@dataclass(frozen=True)
class Dataset:
    """One row of sheet 00 — spec §3."""

    sheet_name: str
    key: str
    headers: tuple[str, ...]
    attributes: tuple[str, ...]

    @property
    def key_field(self) -> str:
        """The first declared header is the key; the rest are measures."""
        return self.headers[0]

    @property
    def measures(self) -> tuple[str, ...]:
        return self.headers[1:]


@dataclass
class Attribute:
    name: str
    value: Any
    is_hypothesis: bool
    row: int

    @property
    def confidence(self) -> Confidence:
        return Confidence.ASSUMED if self.is_hypothesis else Confidence.CONFIRMED


@dataclass
class Hypothesis:
    """Generated from an ``H_`` marker or raised by the tool — spec §5.4."""

    id: str
    dataset_key: str
    attribute: str
    value: Any
    confidence: Confidence
    source: str          # "H_ marker" | "tool"
    note: str = ""

    @property
    def status(self) -> str:
        return "open" if self.confidence is not Confidence.CONFIRMED else "confirmed"


@dataclass
class Record:
    """One extracted record, with provenance — spec §7."""

    source_ref: str                 # "9" row-wise, "D" transposed
    values: dict[str, Any]
    confidence: Confidence | None = None   # populated for transposed blocks only


@dataclass
class Block:
    """A resolved and extracted block — spec §6."""

    dataset: Dataset
    sheet_name: str
    index: int
    orientation: Orientation
    header_ref: str                 # "8" row-wise, "C" transposed
    info_ref: str                   # "L" row-wise, "16" transposed
    address_map: dict[str, str]     # label -> column letter | row number (as str)
    records: list[Record] = field(default_factory=list)
    excluded_records: list[Record] = field(default_factory=list)
    attributes: dict[str, Attribute] = field(default_factory=dict)
    hypotheses: list[Hypothesis] = field(default_factory=list)
    candidates: int = 0
    excluded: int = 0
    unextracted: list[str] = field(default_factory=list)

    @property
    def confidence(self) -> Confidence:
        """Worst confidence of any hypothesis touching the dataset — spec §5.3."""
        return Confidence.worst(h.confidence for h in self.hypotheses)

    @property
    def provenance_label(self) -> str:
        return "Source row" if self.orientation is Orientation.ROW_WISE else "Source column"

    def totals(self) -> dict[str, float]:
        out = {}
        for m in self.dataset.measures:
            vals = [r.values.get(m) for r in self.records]
            out[m] = sum(v for v in vals if isinstance(v, (int, float)))
        return out


class ExtractionError(Exception):
    """A condition the spec requires to be fatal — spec §7.2, §8, §12."""
