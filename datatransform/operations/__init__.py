"""Step 2 as a declared pipeline. Specification_v1.md §9.2.1.

Step 2 used to be one function and a dataclass with thirteen optional fields, and every
new mechanic added one more of each. Two of its rules were not in the declaration at all
but in the *order the code happened to run them*: the split has to come before the
identity, because the identity has nothing to reconcile until the buckets exist, and the
change columns have to come after the sort, because "the year before" is meaningless in
an unsorted list. Both were found by being got wrong.

So a dataset now declares an **ordered list of operations**::

    STEP2["06 EQ Aggs"] = Pipeline(
        Split(total="Total"),          # first: the identity needs the buckets
        Identity(total="Total"),
        Complete(key="Zone", catalogue_attribute="Zone scheme"),
        SortBy(("Zone",)),
        Cumulative("Cumulative exposure %", "Total"),
    )

Each operation reads a :class:`Context` and returns the same one, richer. The order is
declared, so it can be read, reviewed and argued about; a new mechanic is a new class
implementing one method, not a fourteenth field.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Protocol, runtime_checkable

from ..model import Block, Confidence, Record


@dataclass
class Column:
    """A step-2 column that is not a declared field — spec §9.2 S3.

    ``formula`` and ``running_of`` are how the *writer* knows to emit a live Excel
    formula rather than a value. Carrying them here rather than letting the writer reach
    back into the spec is what keeps the writer from knowing what a "calculation" is.
    """

    name: str
    values: list
    number_format: str
    formula: str | None = None        # per row, fields in braces: "{Losses}/{Premium}"
    running_of: str | None = None     # a running share of this field's total — S17


@dataclass
class Figure:
    """A computed block-level figure — spec §9.2.1 S12."""

    name: str
    value: float | None
    detail: str
    number_format: str
    note: str = ""


@dataclass
class AggregateTable:
    """A second step-2 table — spec §9.2.1 S14."""

    title: str
    group_by: str
    measures: tuple[str, ...]
    rows: list[tuple]                       # (group value, {measure: total})
    zero_filled: list[str] = field(default_factory=list)
    note: str = ""

    def totals(self) -> dict[str, float]:
        return {m: sum(v.get(m, 0.0) for _, v in self.rows) for m in self.measures}


@dataclass
class Context:
    """What an operation reads and what it may add to.

    Deliberately mutable and passed along: the alternative is every operation returning a
    six-tuple, which is what the old ``_split`` did and what made it hard to read.
    """

    block: Block
    records: list[Record]
    nomenclature: object = None
    blocks: object = None

    derived_fields: tuple[str, ...] = ()     # declared fields step 2 worked out
    assumed_fields: tuple[str, ...] = ()     # fields resting on a declared ratio — §2.6
    derived: list[Column] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    figures: list[Figure] = field(default_factory=list)
    aggregates: list[AggregateTable] = field(default_factory=list)
    level_finding: object | None = None   # two views of the same book — §2.6
    flags: dict = field(default_factory=dict)   # what one operation tells the next

    @property
    def actual_year(self):
        return getattr(self.nomenclature, "actual_year", None)

    @property
    def period_order(self) -> dict:
        return getattr(self.nomenclature, "period_order", None) or {}

    def known(self) -> set[str]:
        """Fields available to later operations: extracted, plus anything derived."""
        return set(self.block.fields) | set(self.derived_fields)

    def derive(self, *names: str) -> None:
        self.derived_fields += tuple(n for n in names if n not in self.derived_fields)


@runtime_checkable
class Operation(Protocol):
    """One step of step 2. Declared in order; each says what it did, in the sheet."""

    def apply(self, ctx: Context) -> None:
        """Read ``ctx``, change it, and append a note saying what was done."""

    def not_summable(self) -> set[str]:
        """Fields this operation makes that are numbers but not quantities — §2.4."""
        return set()


@dataclass(frozen=True)
class Pipeline:
    """A dataset's step-2 mechanics, in the order they run."""

    operations: tuple[Operation, ...]
    not_summable_fields: tuple[str, ...] = ()

    def __init__(self, *operations: Operation, not_summable: tuple[str, ...] = ()):
        object.__setattr__(self, "operations", tuple(operations))
        object.__setattr__(self, "not_summable_fields", tuple(not_summable))

    def excluded_measures(self) -> set[str]:
        """Numbers that must never be totalled — band bounds, rates. Spec §2.4, §2.7."""
        out = set(self.not_summable_fields)
        for operation in self.operations:
            out |= operation.not_summable()
        return out

    def run(self, ctx: Context) -> Context:
        for operation in self.operations:
            operation.apply(ctx)
        return ctx

    def of_type(self, kind):
        return [o for o in self.operations if isinstance(o, kind)]


@dataclass
class Step2Result:
    """What step 2 produced, and how it is to be written."""

    block: Block
    pipeline: Pipeline
    records: list[Record]
    derived: list[Column] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    figures: list[Figure] = field(default_factory=list)
    aggregates: list[AggregateTable] = field(default_factory=list)
    derived_fields: tuple[str, ...] = ()
    assumed_fields: tuple[str, ...] = ()
    level_finding: object | None = None

    @property
    def computed(self) -> dict[str, list]:
        """The derived columns by name — what the writer fills in."""
        return {c.name: c.values for c in self.derived}

    @property
    def confidence(self) -> Confidence:
        """Step 1's confidence, but never better than the ratios it was expanded with.

        A block extracted cleanly is ``Confirmed``. If three of its five columns exist
        only because a declared split was multiplied onto a reported figure, the block as
        a whole is no longer a reading, and saying so is the entire point of carrying
        confidence at all — spec §5.3.
        """
        if not self.assumed_fields:
            return self.block.confidence
        return Confidence.worst((self.block.confidence, Confidence.ASSUMED))

    @property
    def declared_columns(self) -> tuple[str, ...]:
        """Exactly the order sheet 00 declares, less any field neither present nor derived."""
        available = set(self.block.fields) | set(self.derived_fields)
        return tuple(h for h in self.block.dataset.headers if h in available)

    @property
    def columns(self) -> tuple[str, ...]:
        """Declared columns, then derived ones — spec §9.2 S3."""
        return self.declared_columns + tuple(c.name for c in self.derived)

    def totals(self) -> dict[str, float]:
        return {
            m: sum(r.values[m] for r in self.records
                   if isinstance(r.values.get(m), (int, float)))
            for m in self.block.measure_fields
        }
