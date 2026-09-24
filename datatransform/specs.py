"""What each dataset's step 2 does, and in what order. Specification_v1.md §9.2.1.

Held in code rather than in sheet 00 because these are *mechanics*: sheet 00 carries the
facts a human knows and a machine cannot infer, and the order in which a total is
reconciled is not one of them.

Each dataset declares a :class:`~datatransform.operations.Pipeline` — an ordered list of
operations. The order is the point. Two of these sequences encode a rule that used to
live only in the order the code happened to run:

* ``Split`` before ``Identity`` — the identity has nothing to reconcile until the
  buckets exist.
* ``SortBy`` before ``Change`` — "the year before" is meaningless in an unsorted list.

Both were found by being got wrong. Written down, they can be read and argued about.
"""

from __future__ import annotations

import re

from .constants import (
    A_OCCURRENCE_FROM,
    FMT_AMOUNT,
    FMT_PERCENT,
    F_CATEGORY,
    F_EPI,
    F_EXPOSURE,
    F_INCURRED,
    F_LOSS_AMOUNT,
    F_PREMIUM,
    F_RATE,
    F_RATE_CHANGE,
    F_TOTAL,
    F_YEAR,
    F_ZONE,
    ROLE_HISTORY,
)
from .operations import Pipeline
from .operations.bounds import BandContinuity, DeriveBounds
from .operations.columns import Calculate, Change, Cumulative
from .operations.order import SortBy
from .operations.reconcile import Complete, Identity
from .operations.split import Split
from .operations.tables import Aggregate, DerivedFigure

SPEC_VERSION = "1"

# Field datatypes are declared in ⟦TYPES⟧ of sheet 00, not here: a field name carries
# one meaning across the workbook, which is what a nomenclature is for (spec §8.4).

STEP2: dict[str, Pipeline] = {
    "01 History": Pipeline(
        SortBy((F_YEAR,)),
        Calculate(
            name="Loss Ratio %",
            expression="{" + F_INCURRED + "}/{" + F_PREMIUM + "}",
            number_format=FMT_PERCENT,
            guard_zero=F_PREMIUM,
        ),
    ),
    "02 EPI": Pipeline(
        SortBy((F_YEAR,)),
        DerivedFigure(
            name="Estimation error",
            numerator="{N} re-est",
            denominator="{N} est",
            field=F_EPI,
            note="how far the cedent's own projection for N has moved; "
                 "the measure of how much to trust N+1",
        ),
        DerivedFigure(
            name="Implied growth",
            numerator="{N+1}",
            denominator="{N} re-est",
            field=F_EPI,
            note="cross-check against rate development and exposure growth",
        ),
    ),
    "03 Large": Pipeline(
        SortBy((F_YEAR, "Date of Loss")),
        Aggregate(
            title="Annual sum of large losses",
            group_by=F_YEAR,
            measures=(F_LOSS_AMOUNT,),
            zero_fill_from=ROLE_HISTORY,
        ),
    ),
    "04 Cat": Pipeline(
        SortBy((F_YEAR, "Event Date")),
        Aggregate(
            title="Annual sum of cat losses",
            group_by=F_YEAR,
            measures=(F_LOSS_AMOUNT,),
            zero_fill_from=ROLE_HISTORY,
            note="on the treaty's year basis — this is the table that ties to 01",
        ),
        Aggregate(
            title="By occurrence year",
            group_by="year(Event Date)",
            measures=(F_LOSS_AMOUNT,),
            group_by_attribute=A_OCCURRENCE_FROM,
            note="informational: an event may fall in several underwriting years, so "
                 "this does not tie to 01 unless the treaty is on an occurrence-year "
                 "basis",
        ),
        Aggregate(
            title="By event",
            group_by="Event ID",
            measures=(F_LOSS_AMOUNT,),
            note="what the event actually cost — the figure a cat layer is priced "
                 "against, and the one no annual row shows",
        ),
    ),
    "05 Profile": Pipeline(
        # Bounds first: the sort depends on them, and they may have to be read off the
        # label. By the lower bound as a *number*, never by the label — natural
        # alphanumeric reads "1,000 – 5,000" as the digits 1 then 000, placing it before
        # "500 – 1,000".
        DeriveBounds(label="Band", lower="Band from", upper="Band to"),
        SortBy(("Band from",), numeric=True),
        Calculate(
            name="Average exposure per risk",
            expression="{Exposure}/{Number of Risks}",
            number_format=FMT_AMOUNT,
            guard_zero="Number of Risks",
        ),
        Calculate(
            name="Rate on exposure ‰",
            expression="{Premium}/{Exposure}*1000",
            number_format="0.000",
            guard_zero="Exposure",
        ),
        Cumulative("Cumulative risks %", "Number of Risks"),
        Cumulative("Cumulative exposure %", F_EXPOSURE),
        Cumulative("Cumulative premium %", "Premium"),
        BandContinuity(),
    ),
    # 06 and 07 are the same dataset zoned differently — one profile of the portfolio by
    # geography, because cat losses correlate spatially. Spec §2.5.
    "06 EQ Aggs": Pipeline(
        Split(total=F_TOTAL),                       # before Identity: it needs the buckets
        Identity(total=F_TOTAL),                    # parts from ⟦AXES⟧: nine, or three
        Complete(key=F_ZONE, catalogue_attribute="Zone scheme"),
        SortBy((F_ZONE,)),
        Cumulative("Cumulative exposure %", F_TOTAL),
    ),
    # 08 is the cedent's own split table — the book, not the zones. A dataset in its own
    # right *and* the source of the ratios 06/07 bridge with, which is why it carries the
    # same identity check: a Total that disagrees with its parts would quietly distort
    # every zone downstream.
    "08 Splits": Pipeline(
        Identity(total=F_TOTAL, parts=("Building", "Content", "BI")),
        SortBy((F_CATEGORY,)),
        Cumulative("Share of the book %", F_TOTAL),
    ),
    # 09 is the underwriter's, not the cedent's. Either column may be the one that
    # arrived — given the rates, the change is worked out; given the change, the rates are
    # not. SortBy first: a change is against the year before, which needs an order.
    "09 Rate": Pipeline(
        SortBy((F_YEAR,)),
        Change(name=F_RATE_CHANGE, field=F_RATE, key=F_YEAR),
    ),
}

# The transposed reference sheet is the same dataset in a different orientation.
STEP2["07 Wind Aggs"] = STEP2["06 EQ Aggs"]
STEP2["01 History_T"] = STEP2["01 History"]
STEP2["02 EPI_T"] = STEP2["02 EPI"]

STEP2_BY_ROLE: dict[str, Pipeline] = {
    "01": STEP2["01 History"],
    "02": STEP2["02 EPI"],
    "03": STEP2["03 Large"],
    "04": STEP2["04 Cat"],
    "05": STEP2["05 Profile"],
    "06": STEP2["06 EQ Aggs"],
    "07": STEP2["06 EQ Aggs"],      # windstorm: the same dataset, a different zoning
    "08": STEP2["08 Splits"],
    "09": STEP2["09 Rate"],
}


def step2_for(dataset_key: str) -> Pipeline | None:
    """Exact key first, then the role — spec §2.3.

    A multi-section pack names its sheets ``01 History Fire``, ``01 History EQ`` and so
    on; they share role ``01`` and therefore the same step-2 mechanics.
    """
    pipeline = STEP2.get(dataset_key)
    if pipeline is not None:
        return pipeline
    match = re.match(r"^\s*(\d+)", str(dataset_key))
    return STEP2_BY_ROLE.get(match.group(1).zfill(2)) if match else None
