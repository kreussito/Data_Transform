"""Step 1 and step 2. Specification_v1.md §9.2, §10."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .model import Block, ExtractionError, Record
from .specs import Step2Spec

FIELD = re.compile(r"\{([^}]+)\}")


@dataclass
class Step2Result:
    block: Block
    spec: Step2Spec
    records: list[Record]
    columns: tuple[str, ...]
    computed: dict[str, list[float | None]] = field(default_factory=dict)
    notes: list[str] = field(default_factory=list)

    def totals(self) -> dict[str, float]:
        out = {}
        for m in self.block.dataset.measures:
            out[m] = sum(
                r.values[m] for r in self.records if isinstance(r.values.get(m), (int, float))
            )
        return out


def _sort_key(record: Record, fields):
    key = []
    for f in fields:
        v = record.values.get(f)
        key.append((v is None, v if v is not None else 0))
    return key


def apply_step2(block: Block, spec: Step2Spec) -> Step2Result:
    missing = [f for f in spec.sort_by if f not in block.dataset.headers]
    if missing:
        raise ExtractionError(f"step 2 sorts by {missing}, which the dataset does not declare")

    records = sorted(block.records, key=lambda r: _sort_key(r, spec.sort_by),
                     reverse=not spec.ascending)

    computed: dict[str, list[float | None]] = {}
    for calc in spec.calculations:
        column = []
        for r in records:
            expr = calc.expression
            guard = r.values.get(calc.guard_zero) if calc.guard_zero else None
            if calc.guard_zero and (guard in (None, 0)):
                column.append(None)
                continue
            try:
                resolved = FIELD.sub(lambda m: repr(float(r.values[m.group(1)])), expr)
                column.append(eval(resolved, {"__builtins__": {}}, {}))  # noqa: S307
            except (KeyError, TypeError, ValueError, ZeroDivisionError):
                column.append(None)
        computed[calc.name] = column

    notes = [
        f"sorted by {', '.join(spec.sort_by)} "
        f"{'ascending' if spec.ascending else 'descending'} (value-preserving)",
        f"columns ordered {' | '.join(spec.column_order)} (value-preserving)",
    ]
    for calc in spec.calculations:
        readable = FIELD.sub(lambda m: m.group(1), calc.expression).replace("/", " / ")
        notes.append(f"calculated {calc.name} = {readable} (value-adding)")

    result = Step2Result(
        block=block, spec=spec, records=records,
        columns=spec.output_columns(), computed=computed, notes=notes,
    )

    # Sorting and reordering cannot move a total; if they do, that is a bug — spec §10 C3.
    before, after = block.totals(), result.totals()
    for measure, total in before.items():
        if abs(total - after[measure]) > 1e-9:
            raise ExtractionError(
                f"step 2 is value-preserving for {measure!r} but the total moved "
                f"from {total} to {after[measure]}"
            )
    return result
