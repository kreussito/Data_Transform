"""Step-2 mechanics, held in code rather than in sheet 00.

Specification_v1.md §3: sheet 00 carries facts a human knows and a machine cannot
infer. Sort order, target column order and derived measures are *mechanics*, so they
live here, versioned with the tool.
"""

from __future__ import annotations

from dataclasses import dataclass, field

SPEC_VERSION = "1"

# Field datatypes are declared in ⟦TYPES⟧ of sheet 00, not here: a field name carries
# one meaning across the workbook, which is what a nomenclature is for (spec §8.4).


@dataclass(frozen=True)
class Calculation:
    """A derived measure. ``expression`` names fields in braces."""

    name: str
    expression: str
    number_format: str = "0.0%"
    guard_zero: str | None = None      # field that must not be zero


@dataclass(frozen=True)
class DerivedFigure:
    """A step-2 figure relating two *records* rather than describing one — spec §9.2.1.

    Written beneath the data, not as a column, because it has no per-row meaning.
    """

    name: str
    numerator: str            # record pattern, e.g. "{N} re-est"
    denominator: str          # record pattern, e.g. "{N} est"
    field: str
    kind: str = "ratio_minus_1"
    number_format: str = "0.0%"
    note: str = ""


@dataclass(frozen=True)
class AggregateSpec:
    """A second step-2 table: group by a key, sum measures — spec §9.2.1 S14.

    ``zero_fill_from`` names a dataset whose year window the table should span, so a
    year with no records shows 0 rather than being absent. A missing year reads as
    "no data"; a zero reads as "nothing happened", and only one of those is true.
    """

    title: str
    group_by: str
    measures: tuple[str, ...]
    zero_fill_from: str | None = None


@dataclass(frozen=True)
class Step2Spec:
    """Step-2 mechanics.

    There is deliberately no ``column_order``: the output order is the order sheet 00
    declares (spec §9.2 S2), so that 00 stays the single source of truth for both what
    is extracted and in what order it appears.
    """

    sort_by: tuple[str, ...]
    ascending: bool = True
    calculations: tuple[Calculation, ...] = field(default_factory=tuple)
    figures: tuple[DerivedFigure, ...] = field(default_factory=tuple)
    aggregate: AggregateSpec | None = None


STEP2: dict[str, Step2Spec] = {
    "01 History": Step2Spec(
        sort_by=("Year",),
        ascending=True,
        calculations=(
            Calculation(
                name="Loss Ratio %",
                expression="{Incurred Losses}/{Premium}",
                number_format="0.0%",
                guard_zero="Premium",
            ),
        ),
    ),
    "02 EPI": Step2Spec(
        sort_by=("Year",),
        ascending=True,
        figures=(
            DerivedFigure(
                name="Estimation error",
                numerator="{N} re-est",
                denominator="{N} est",
                field="EPI",
                note="how far the cedent's own projection for N has moved; "
                     "the measure of how much to trust N+1",
            ),
            DerivedFigure(
                name="Implied growth",
                numerator="{N+1}",
                denominator="{N} re-est",
                field="EPI",
                note="cross-check against rate development and exposure growth",
            ),
        ),
    ),
    "03 Large": Step2Spec(
        sort_by=("Year", "Date of Loss"),
        ascending=True,
        aggregate=AggregateSpec(
            title="Annual sum of large losses",
            group_by="Year",
            measures=("Loss amount",),
            zero_fill_from="01 History",
        ),
    ),
}

# The transposed reference sheet is the same dataset in a different orientation.
STEP2["01 History_T"] = STEP2["01 History"]
STEP2["02 EPI_T"] = STEP2["02 EPI"]


def step2_for(dataset_key: str) -> Step2Spec | None:
    return STEP2.get(dataset_key)
