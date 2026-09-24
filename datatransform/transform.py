"""Step 1 and step 2. Specification_v1.md §9.2, §10.

Step 1 is an extraction and lives in :mod:`datatransform.extract`. What is left here is
step 2: running a dataset's declared :class:`~datatransform.operations.Pipeline` over the
records step 1 produced, and handing the result to the writer.

The mechanics themselves are in :mod:`datatransform.operations`, one class per operation.
This module is the façade — it exists so that "apply step 2" is one call, and so that the
names the rest of the tool has always imported from here keep working.
"""

from __future__ import annotations

from .model import Block, Confidence, ExtractionError, Hypothesis, Record  # noqa: F401
from .operations import AggregateTable, Column, Context, Figure, Pipeline, Step2Result
from .operations.order import natural_key as _natural_key
from .operations.order import period_key as _period_key
from .operations.order import sort_key as _sort_key
from .operations.tables import _aggregate  # noqa: F401  (tests reach for it)

__all__ = [
    "AggregateTable", "Column", "Context", "Figure", "Pipeline", "Step2Result",
    "apply_step2", "_natural_key", "_period_key", "_sort_key",
]


def apply_step2(block: Block, pipeline: Pipeline, nomenclature=None,
                blocks=None) -> Step2Result:
    """Run the dataset's declared operations, in the order they are declared.

    Every operation reads the context and adds to it — records, columns, notes, tables,
    findings. Nothing here decides *what* happens; that is in
    :data:`datatransform.specs.STEP2`, where it can be read and argued about.
    """
    ctx = Context(block=block, records=list(block.records),
                  nomenclature=nomenclature, blocks=blocks)
    pipeline.run(ctx)

    # An invariant, not an operation: the columns are always shown in the order sheet 00
    # declares, and the note has to come after everything that can add a field. Leaving
    # that to each pipeline was one more way to get an ordering wrong.
    shown = [h for h in block.dataset.headers if h in ctx.known()]
    ctx.notes.append(
        f"columns ordered as declared in sheet 00: {' | '.join(shown)} (value-preserving)")

    result = Step2Result(
        block=block,
        pipeline=pipeline,
        records=ctx.records,
        derived=ctx.derived,
        notes=ctx.notes,
        figures=ctx.figures,
        aggregates=ctx.aggregates,
        derived_fields=ctx.derived_fields,
        assumed_fields=ctx.assumed_fields,
        level_finding=ctx.level_finding,
    )
    _check_invariants(block, result)
    return result


def _check_invariants(block: Block, result: Step2Result) -> None:
    """What must hold whatever the pipeline did — spec §9.2 S6, S15.

    Deliberately **not** operations. An operation is a choice a dataset makes; these two
    are the promise step 2 gives about every choice, so they run last and always.
    """
    # Every grouping of the same records must reach the same total — S15.
    for table in result.aggregates:
        for measure, total in table.totals().items():
            if abs(total - result.totals().get(measure, 0.0)) > 1e-9:
                raise ExtractionError(
                    f"{table.title!r}: the aggregate total for {measure!r} is {total}, "
                    f"but the detail totals {result.totals().get(measure)} — "
                    "grouping lost records"
                )

    # Sorting and reordering cannot move a total; if they do, that is a bug — S6.
    before, after = block.totals(), result.totals()
    for measure, total in before.items():
        if abs(total - after[measure]) > 1e-9:
            raise ExtractionError(
                f"step 2 is value-preserving for {measure!r} but the total moved "
                f"from {total} to {after[measure]}"
            )
