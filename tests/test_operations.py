"""Step 2 as a declared pipeline — specification_v1.md §9.2.1.

The point of the pipeline is that the *order* is declared rather than implicit. Two rules
used to live only in the sequence the code happened to run, and both were found by being
got wrong. These tests hold them where they can now be read.
"""

from __future__ import annotations

import pytest

from datatransform.operations import Column, Context, Pipeline
from datatransform.operations.columns import Calculate, Change, Cumulative
from datatransform.operations.order import SortBy
from datatransform.operations.reconcile import Complete, Identity
from datatransform.operations.split import Split
from datatransform.operations.tables import Aggregate, DerivedFigure
from datatransform.specs import STEP2, STEP2_BY_ROLE, step2_for


def _index(pipeline, kind) -> int:
    return next(i for i, op in enumerate(pipeline.operations) if isinstance(op, kind))


def _has(pipeline, kind) -> bool:
    return any(isinstance(op, kind) for op in pipeline.operations)


# ───────────────────────────── the orders that used to be implicit

def test_the_split_runs_before_the_identity():
    """The identity has nothing to reconcile until the buckets exist."""
    for key in ("06 EQ Aggs", "07 Wind Aggs"):
        pipeline = STEP2[key]
        assert _index(pipeline, Split) < _index(pipeline, Identity), key


def test_the_sort_runs_before_any_change():
    """'The year before' is meaningless in an unsorted list."""
    pipeline = STEP2["09 Rate"]
    assert _index(pipeline, SortBy) < _index(pipeline, Change)


def test_the_bounds_are_derived_before_the_sort_that_uses_them():
    """05 sorts on a bound that may have to be read off the label first — §2.4."""
    from datatransform.operations.bounds import DeriveBounds

    pipeline = STEP2["05 Profile"]
    assert _index(pipeline, DeriveBounds) < _index(pipeline, SortBy)
    assert pipeline.operations[_index(pipeline, SortBy)].numeric


def test_completing_the_zone_list_happens_after_the_split():
    """A zone added as 0 must carry zeros in the split buckets too."""
    pipeline = STEP2["06 EQ Aggs"]
    assert _index(pipeline, Split) < _index(pipeline, Complete)


# ─────────────────────────────── an operation says what it makes

def test_an_operation_declares_what_must_never_be_totalled():
    """Five years' rates do not add up to a rate — §2.4's rule, stated by the operation
    that creates the column rather than repeated in the dataset's declaration."""
    assert Change(name="Rate change", field="Rate").not_summable() == {"Rate change",
                                                                      "Rate"}
    assert STEP2["09 Rate"].excluded_measures() == {"Rate", "Rate change"}


def test_the_band_bounds_exclude_themselves():
    assert STEP2["05 Profile"].excluded_measures() == {"Band from", "Band to"}


def test_a_dataset_with_nothing_to_exclude_excludes_nothing():
    assert STEP2["01 History"].excluded_measures() == set()


# ─────────────────────────────────────── every dataset is reachable

def test_every_role_resolves_to_a_pipeline():
    for role in ("01", "02", "03", "04", "05", "06", "07", "08", "09"):
        assert isinstance(STEP2_BY_ROLE[role], Pipeline)
    assert isinstance(step2_for("01 History Fire"), Pipeline)      # by role
    assert step2_for("99 Nothing") is None


def test_every_pipeline_sorts():
    """Nothing downstream means anything in an arbitrary order."""
    for key, pipeline in STEP2.items():
        assert _has(pipeline, SortBy), key


# ──────────────────────────────── an operation runs on its own

def test_an_operation_needs_no_workbook():
    """The point of one class per operation: it can be tested without a spreadsheet."""
    from datatransform.model import Dataset, FieldType, Orientation, Record
    from datatransform.model import Block

    dataset = Dataset("x", "01 History", ("Year", "Premium"), ())
    block = Block(dataset=dataset, sheet_name="x", index=1,
                  orientation=Orientation.ROW_WISE, header_ref="1", info_ref="B",
                  address_map={"Year": "A", "Premium": "B"},
                  field_types={"Year": FieldType.TEXT, "Premium": FieldType.NUMBER},
                  records=[Record("2", {"Year": "2024", "Premium": 10.0}),
                           Record("3", {"Year": "2023", "Premium": 20.0})])
    ctx = Context(block=block, records=list(block.records))

    SortBy(("Year",)).apply(ctx)
    assert [r.values["Year"] for r in ctx.records] == ["2023", "2024"]
    assert "sorted by Year ascending" in ctx.notes[0]

    Cumulative("Share %", "Premium").apply(ctx)
    assert [c.name for c in ctx.derived] == ["Share %"]
    assert ctx.derived[0].values == [pytest.approx(2 / 3), pytest.approx(1.0)]
    assert ctx.derived[0].running_of == "Premium"


def test_a_calculation_carries_its_formula_for_the_writer():
    """The writer emits a live formula; it learns how from the column, not from the spec."""
    calc = Calculate(name="Loss Ratio %", expression="{A}/{B}", guard_zero="B")
    assert isinstance(Column("x", [], "0.0%"), Column)
    assert calc.expression == "{A}/{B}"


def test_the_pipeline_is_readable_as_a_list():
    """What a dataset does in step 2 should be answerable by reading five lines."""
    names = [type(op).__name__ for op in STEP2["06 EQ Aggs"].operations]
    assert names == ["Split", "Identity", "Complete", "SortBy", "Cumulative"]

    names = [type(op).__name__ for op in STEP2["04 Cat"].operations]
    assert names == ["SortBy", "Aggregate", "Aggregate", "Aggregate"]
    assert not _has(STEP2["04 Cat"], DerivedFigure)
