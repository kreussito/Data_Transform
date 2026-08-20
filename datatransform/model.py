"""Data model — specification_v1.md §4, §5, §6."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any

from .constants import FAILED, FMT_AMOUNT, FMT_DATE, FMT_TEXT


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
    DATE = "date"

    @property
    def number_format(self) -> str:
        return {
            FieldType.TEXT: FMT_TEXT,
            FieldType.NUMBER: FMT_AMOUNT,
            FieldType.DATE: FMT_DATE,
        }[self]

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
    optional: tuple[str, ...] = ()      # declared in 00 as "<name> (optional)"

    def is_optional(self, header: str) -> bool:
        return header in self.optional

    @property
    def key_field(self) -> str:
        """The first declared header identifies the record."""
        return self.headers[0]

    @property
    def role(self) -> str:
        """The leading number of the key — spec §2.3.

        ``01 History Fire`` and ``01 History EQ`` both play role ``01``: they are the
        same dataset for different sections of the treaty.
        """
        import re as _re

        match = _re.match(r"^\s*(\d+)", self.key)
        return match.group(1).zfill(2) if match else self.key


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
    section: str | None = None      # which section of the treaty, if declared
    candidates: int = 0
    excluded: int = 0
    claimed_elsewhere: int = 0            # rows another block on this sheet extracts
    claimed_by: list[str] = field(default_factory=list)
    unextracted: list[str] = field(default_factory=list)
    display_formats: dict[str, str] = field(default_factory=dict)

    @property
    def confidence(self) -> Confidence:
        """Worst confidence of any hypothesis touching the dataset — spec §5.3."""
        return Confidence.worst(h.confidence for h in self.hypotheses)

    @property
    def provenance_label(self) -> str:
        return "Source row" if self.orientation is Orientation.ROW_WISE else "Source column"

    @property
    def fields(self) -> tuple[str, ...]:
        """Declared headers that this block actually resolved — spec §8 F4."""
        return tuple(h for h in self.dataset.headers if h in self.address_map)

    @property
    def numeric_fields(self) -> tuple[str, ...]:
        """Every field read as a number — spec §8.4."""
        return tuple(h for h in self.fields if self.field_types.get(h) is FieldType.NUMBER)

    @property
    def measure_fields(self) -> tuple[str, ...]:
        """The numeric fields that are **business measures**, and so are summed.

        A band bound is a number, but summing the lower edges of a risk profile produces
        a figure that means nothing and would sit in the control row inviting a reader to
        interpret it. Bounds order and compare records; they are not quantities of
        anything — spec §2.4. A rate is the same kind of number: the sum of five years'
        rates is not a rate, and a total of five rate *changes* is not a change.
        """
        from .specs import step2_for

        spec = step2_for(self.dataset.key)
        bounds = getattr(spec, "bounds", None) if spec else None
        excluded = {bounds.lower, bounds.upper} if bounds else set()
        excluded |= set(getattr(spec, "not_summable", ()) if spec else ())
        return tuple(f for f in self.numeric_fields if f not in excluded)

    def number_format(self, label: str) -> str:
        """How to display a figure — as the source displayed it, or as step 2 says.

        The declared type gives a sensible default, but only the source knows whether a
        number is an amount, a percentage or a rate per mille. Carrying its format across
        is not interpretation; writing 0.00118 as ``0`` would be. A column step 2 derived
        has no source cell, so the rule that derived it says how it reads.
        """
        declared = self.field_types.get(label, FieldType.NUMBER)
        if declared is FieldType.NUMBER and label in self.display_formats:
            return self.display_formats[label]
        return declared.number_format

    def totals(self) -> dict[str, float]:
        return {
            m: sum(r.values[m] for r in self.records
                   if isinstance(r.values.get(m), (int, float)))
            for m in self.measure_fields
        }


@dataclass(frozen=True)
class Section:
    """One row of ⟦SECTIONS⟧ — spec §2.3.

    A treaty is one or more sections. Fire is *per risk* and carries large losses;
    Earthquake and Windstorm are *cat* and carry event losses instead. Which datasets
    a section expects is declared, so an absent sheet is structure rather than a gap.
    """

    name: str
    kind: str
    roles: tuple[str, ...] = ()

    def expects(self, role: str) -> bool:
        return not self.roles or role in self.roles


@dataclass(frozen=True)
class Axis:
    """One row of ⟦AXES⟧ — a dimension a dataset is split along, spec §2.6.

    The buckets of a dataset are the **product** of its axes, which is what makes one
    mechanism serve four different tables: earthquake and fire are occupancy × cover
    (nine), engineering is occupancy × Projects/Renewables (six), and windstorm declares
    the occupancy axis alone (three).
    """

    dataset: str
    name: str
    categories: tuple[str, ...]


@dataclass(frozen=True)
class SplitRule:
    """One row of ⟦SPLITS⟧ — a declared ratio, spec §2.6.

    Two jobs in one shape. With ``source_category`` empty the rule distributes a whole
    axis the cedent never reported. With it set, the rule re-splits a category that
    arrived merged — ``Commercial`` covering both commercial and industrial risks.

    ``source`` is not decoration. A ratio taken from the cedent's own prior submission
    and one borrowed from another book are both assumptions, but they are not equally
    good ones, and the reader has to be able to tell them apart.
    """

    dataset: str
    axis: str
    source_category: str
    category: str
    share: float
    source: str = ""

    def applies_to(self, dataset_key: str) -> bool:
        return self.dataset == "*" or self.dataset.casefold() == dataset_key.casefold()


@dataclass(frozen=True)
class Reference:
    """One side of a rule — spec §10.1.

    Three forms, distinguished by shape alone:

    ==================================== ==========================================
    ``01 History.Premium@{N}``           a field on the record matching ``{N}``
    ``01 History.Year basis``            an *attribute* of the sheet — no record
    ``SUM(03 Large.Loss amount@{Y})``    the field summed over every matching record
    ==================================== ==========================================

    ``{N}`` and ``{N+1}`` resolve from ⟦GLOBAL⟧; ``{Y}`` is a wildcard that expands
    the rule once per year present in the data.
    """

    dataset_key: str
    name: str
    record: str | None = None
    aggregate: str | None = None

    @property
    def is_attribute(self) -> bool:
        return self.record is None

    def with_record(self, record: str) -> "Reference":
        return Reference(self.dataset_key, self.name, record, self.aggregate)

    def render(self) -> str:
        inner = f"{self.dataset_key}.{self.name}"
        if self.record is not None:
            inner += f"@{self.record}"
        return f"{self.aggregate}({inner})" if self.aggregate else inner


AGGREGATES = ("SUM",)


def parse_reference(ref: str) -> Reference:
    text = str(ref).strip()

    aggregate = None
    for name in AGGREGATES:
        if text.upper().startswith(name + "(") and text.endswith(")"):
            aggregate, text = name, text[len(name) + 1:-1].strip()
            break

    head, sep, record = text.partition("@")
    key, dot, field = head.rpartition(".")
    if not key or not dot or not field:
        raise ExtractionError(
            f"sheet 00 ⟦RULES⟧: {ref!r} is not of the form <dataset key>.<field>@<record>, "
            f"<dataset key>.<attribute>, or SUM(<dataset key>.<field>@<record>)"
        )
    if sep and not record.strip():
        raise ExtractionError(f"sheet 00 ⟦RULES⟧: {ref!r} has an empty record selector")
    if aggregate and not sep:
        raise ExtractionError(f"sheet 00 ⟦RULES⟧: {ref!r} aggregates but names no record")

    return Reference(key.strip(), field.strip(),
                     record.strip() if sep else None, aggregate)


@dataclass(frozen=True)
class Rule:
    """One row of ⟦RULES⟧ in sheet 00 — spec §10.1."""

    id: str
    left: str
    relation: str
    right: str
    tolerance: float
    severity: str
    note: str = ""
    scope: str = ""            # "" | "all" | a section kind | a section name

    @property
    def left_ref(self) -> Reference:
        return parse_reference(self.left)

    @property
    def right_ref(self) -> Reference:
        return parse_reference(self.right)

    @staticmethod
    def parse_ref(ref: str) -> tuple[str, str, str]:
        """Backwards-compatible tuple view."""
        r = parse_reference(ref)
        return r.dataset_key, r.name, r.record or ""


@dataclass
class RuleResult:
    """The outcome of evaluating a Rule — spec §10.1."""

    rule: Rule
    status: str                # passed | failed | skipped | not applicable
    detail: str
    left_value: float | None = None
    right_value: float | None = None
    year: int | None = None      # set when a {Y} rule was expanded
    section: str | None = None   # set when the rule was expanded over sections

    @property
    def ok(self) -> bool:
        return self.status != FAILED

    @property
    def label(self) -> str:
        parts = [self.rule.id]
        if self.section:
            parts.append(self.section)
        if self.year is not None:
            parts.append(str(self.year))
        return "/".join(parts)


class ExtractionError(Exception):
    """A condition the spec requires to be fatal — spec §7.2, §8, §12."""
