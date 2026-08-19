"""Step-2 mechanics, held in code rather than in sheet 00.

Specification_v1.md §3: sheet 00 carries facts a human knows and a machine cannot
infer. Sort order, target column order and derived measures are *mechanics*, so they
live here, versioned with the tool.
"""

from __future__ import annotations

import re
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

    ``group_by_attribute`` names an attribute that may redirect the grouping — spec
    §9.2.1 S15. Which date defines the occurrence year is the cedent's convention, not
    the tool's, so ``Occurrence year from = Event End Date`` changes the table without
    changing the code.
    """

    title: str
    group_by: str                     # a field, or ``year(<date field>)``
    measures: tuple[str, ...]
    zero_fill_from: str | None = None
    group_by_attribute: str | None = None
    note: str = ""


@dataclass(frozen=True)
class Bounds:
    """Numeric bounds that may have to be read off a label — spec §2.4, §9.2.1 S16.

    Where the source declares the two columns they are extracted and this does nothing.
    Where it does not, step 2 reads them off ``label`` and says so — interpretation
    belongs to step 2, never to step 1.
    """

    label: str
    lower: str
    upper: str


@dataclass(frozen=True)
class Cumulative:
    """A running share of the column's total — spec §9.2.1 S17.

    What a profile exists to answer is "where does the book sit?", and that is a
    cumulative question: no single band answers it.
    """

    name: str
    field: str
    number_format: str = "0.0%"


@dataclass(frozen=True)
class Complete:
    """Every member of a declared catalogue must appear — spec §9.2.1 S18.

    A cat aggregate lists only the zones the cedent has exposure in. The zones with
    none are exactly the ones worth seeing, so the missing ones are added with 0.
    Adding zeros cannot move a total, so this stays value-preserving.

    ``catalogue_attribute`` names the attribute that selects which list applies —
    ``Zone scheme = Mexico EQ`` and ``Mexico Wind`` are different zonings of the same
    country and must not be mixed.
    """

    key: str
    catalogue_attribute: str


@dataclass(frozen=True)
class Identity:
    """A declared field that equals the sum of others — spec §9.2.1 S19.

    Three shapes, like the band bounds:

    * both sides supplied — **checked**, and a disagreement is a finding
    * the total absent — **derived**, and said so
    * the parts absent — the total stands alone, and the gap is named

    ``parts`` empty means *the dataset's buckets*, taken from ⟦AXES⟧ — so a dataset whose
    axes change does not need its identity restated.
    """

    total: str
    parts: tuple[str, ...] = ()


@dataclass(frozen=True)
class Split:
    """Expand what arrived into the dataset's full bucket set — spec §2.6, S20.

    The buckets are the product of the declared axes. A cedent reports at whatever level
    it keeps its book: the whole total, one axis, or the finished grid. Step 2 multiplies
    the missing axes onto what came, using ratios that are themselves declared and
    sourced.

    ``total`` names the field to start from when nothing but a grand figure arrived.
    """

    total: str = "Total"


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
    aggregates: tuple[AggregateSpec, ...] = field(default_factory=tuple)
    cumulative: tuple[Cumulative, ...] = field(default_factory=tuple)
    bounds: Bounds | None = None
    numeric_sort: bool = False       # sort on the number, not the label
    complete: Complete | None = None
    identity: Identity | None = None
    split: Split | None = None


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
    "04 Cat": Step2Spec(
        sort_by=("Year", "Event Date"),
        ascending=True,
        aggregates=(
            AggregateSpec(
                title="Annual sum of cat losses",
                group_by="Year",
                measures=("Loss amount",),
                zero_fill_from="01",
                note="on the treaty's year basis — this is the table that ties to 01",
            ),
            AggregateSpec(
                title="By occurrence year",
                group_by="year(Event Date)",
                measures=("Loss amount",),
                group_by_attribute="Occurrence year from",
                note="informational: an event may fall in several underwriting years, "
                     "so this does not tie to 01 unless the treaty is on an "
                     "occurrence-year basis",
            ),
            AggregateSpec(
                title="By event",
                group_by="Event ID",
                measures=("Loss amount",),
                note="what the event actually cost — the figure a cat layer is priced "
                     "against, and the one no annual row shows",
            ),
        ),
    ),
    "05 Profile": Step2Spec(
        # By the lower bound as a number, never by the label: natural alphanumeric reads
        # "1,000 – 5,000" as the digits 1 then 000, placing it before "500 – 1,000".
        sort_by=("Band from",),
        ascending=True,
        numeric_sort=True,
        bounds=Bounds(label="Band", lower="Band from", upper="Band to"),
        calculations=(
            Calculation(
                name="Average exposure per risk",
                expression="{Exposure}/{Number of Risks}",
                number_format="#,##0",
                guard_zero="Number of Risks",
            ),
            Calculation(
                name="Rate on exposure ‰",
                expression="{Premium}/{Exposure}*1000",
                number_format="0.000",
                guard_zero="Exposure",
            ),
        ),
        cumulative=(
            Cumulative("Cumulative risks %", "Number of Risks"),
            Cumulative("Cumulative exposure %", "Exposure"),
            Cumulative("Cumulative premium %", "Premium"),
        ),
    ),
    # 06 and 07 are the same dataset zoned differently — one profile of the portfolio
    # by geography, because cat losses correlate spatially. Spec §2.5.
    "06 EQ Aggs": Step2Spec(
        sort_by=("Zone",),
        ascending=True,
        split=Split(total="Total"),
        complete=Complete(key="Zone", catalogue_attribute="Zone scheme"),
        # The parts are the buckets from ⟦AXES⟧: nine for earthquake, three for wind.
        identity=Identity(total="Total"),
        cumulative=(Cumulative("Cumulative exposure %", "Total"),),
    ),
    "03 Large": Step2Spec(
        sort_by=("Year", "Date of Loss"),
        ascending=True,
        aggregates=(
            AggregateSpec(
                title="Annual sum of large losses",
                group_by="Year",
                measures=("Loss amount",),
                zero_fill_from="01",
            ),
        ),
    ),
}

# The transposed reference sheet is the same dataset in a different orientation.
STEP2["07 Wind Aggs"] = STEP2["06 EQ Aggs"]
STEP2["01 History_T"] = STEP2["01 History"]
STEP2["02 EPI_T"] = STEP2["02 EPI"]


def step2_for(dataset_key: str) -> Step2Spec | None:
    """Exact key first, then the role — spec §2.3.

    A multi-section pack names its sheets ``01 History Fire``, ``01 History EQ`` and so
    on; they share role ``01`` and therefore the same step-2 mechanics.
    """
    spec = STEP2.get(dataset_key)
    if spec is not None:
        return spec
    match = re.match(r"^\s*(\d+)", str(dataset_key))
    return STEP2_BY_ROLE.get(match.group(1).zfill(2)) if match else None


STEP2_BY_ROLE: dict[str, Step2Spec] = {
    "01": STEP2["01 History"],
    "02": STEP2["02 EPI"],
    "03": STEP2["03 Large"],
    "04": STEP2["04 Cat"],
    "05": STEP2["05 Profile"],
    "06": STEP2["06 EQ Aggs"],
    "07": STEP2["06 EQ Aggs"],      # windstorm: the same dataset, a different zoning
}
