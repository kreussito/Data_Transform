"""Data model — specification_v1.md §4, §5, §6."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


class Orientation(str, Enum):
    ROW_WISE = "row-wise"
    TRANSPOSED = "transposed"


class FieldType(str, Enum):
    """How a field is read — spec §8.4.

    Declared per field rather than inferred from position: on ``03. Large Losses``,
    ``Loss Date`` and ``Claim Reference`` sit where a measure would and must never
    be summed.
    """

    TEXT = "text"
    NUMBER = "number"

    @property
    def number_format(self) -> str:
        return "@" if self is FieldType.TEXT else "#,##0"

    @classmethod
    def parse(cls, text: str) -> "FieldType":
        try:
            return cls(str(text).strip().casefold())
        except ValueError:
            raise ExtractionError(
                f"sheet 00 ⟦TYPES⟧: {text!r} is not a known datatype "
                f"({', '.join(t.value for t in cls)})"
            ) from None


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
        """The first declared header identifies the record."""
        return self.headers[0]


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
    field_types: dict[str, FieldType] = field(default_factory=dict)
    records: list[Record] = field(default_factory=list)
    excluded_records: list[Record] = field(default_factory=list)
    attributes: dict[str, Attribute] = field(default_factory=dict)
    hypotheses: list[Hypothesis] = field(default_factory=list)
    coercions: list = field(default_factory=list)
    crosschecks: list = field(default_factory=list)
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

    @property
    def numeric_fields(self) -> tuple[str, ...]:
        """Only these are summed — spec §8.4."""
        return tuple(
            h for h in self.dataset.headers
            if self.field_types.get(h) is FieldType.NUMBER
        )

    def number_format(self, label: str) -> str:
        return self.field_types.get(label, FieldType.NUMBER).number_format

    def totals(self) -> dict[str, float]:
        return {
            m: sum(r.values[m] for r in self.records
                   if isinstance(r.values.get(m), (int, float)))
            for m in self.numeric_fields
        }


@dataclass(frozen=True)
class Rule:
    """One row of ⟦RULES⟧ in sheet 00 — spec §10.1.

    ``left`` and ``right`` are ``<dataset key>.<field>@<record>`` references; ``{N}``
    and ``{N+1}`` resolve from ⟦GLOBAL⟧'s ``Actual year``.
    """

    id: str
    left: str
    relation: str
    right: str
    tolerance: float
    severity: str
    note: str = ""

    @staticmethod
    def parse_ref(ref: str) -> tuple[str, str, str]:
        head, _, record = ref.partition("@")
        key, _, field = head.rpartition(".")
        if not key or not field or not record:
            raise ExtractionError(
                f"sheet 00 ⟦RULES⟧: {ref!r} is not of the form "
                f"<dataset key>.<field>@<record>"
            )
        return key.strip(), field.strip(), record.strip()


@dataclass
class RuleResult:
    """The outcome of evaluating a Rule — spec §10.1."""

    rule: Rule
    status: str                      # passed | failed | skipped
    detail: str
    left_value: float | None = None
    right_value: float | None = None

    @property
    def ok(self) -> bool:
        return self.status != "failed"


class ExtractionError(Exception):
    """A condition the spec requires to be fatal — spec §7.2, §8, §12."""
