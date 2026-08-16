"""Step 1 and step 2. Specification_v1.md §9.2, §10."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .model import Block, ExtractionError, Record
from .specs import AggregateSpec, Step2Spec

FIELD = re.compile(r"\{([^}]+)\}")


@dataclass
class Figure:
    """A computed block-level figure — spec §9.2.1."""

    name: str
    value: float | None
    detail: str
    number_format: str
    note: str = ""


@dataclass
class AggregateTable:
    """The second step-2 table — spec §9.2.1 S14."""

    title: str
    group_by: str
    measures: tuple[str, ...]
    rows: list[tuple]                       # (group value, {measure: total})
    zero_filled: list[str] = field(default_factory=list)

    def totals(self) -> dict[str, float]:
        return {m: sum(values.get(m, 0.0) for _, values in self.rows) for m in self.measures}


@dataclass
class Step2Result:
    block: Block
    spec: Step2Spec
    records: list[Record]
    computed: dict[str, list[float | None]] = field(default_factory=dict)
    notes: list[str] = field(default_factory=list)
    figures: list[Figure] = field(default_factory=list)
    aggregate: AggregateTable | None = None

    @property
    def declared_columns(self) -> tuple[str, ...]:
        """Exactly the order sheet 00 declares — spec §9.2 S2."""
        return tuple(self.block.dataset.headers)

    @property
    def columns(self) -> tuple[str, ...]:
        """Declared columns, then derived ones — spec §9.2 S3."""
        return self.declared_columns + tuple(c.name for c in self.spec.calculations)

    def totals(self) -> dict[str, float]:
        out = {}
        for m in self.block.numeric_fields:
            out[m] = sum(
                r.values[m] for r in self.records if isinstance(r.values.get(m), (int, float))
            )
        return out


DIGITS = re.compile(r"(\d+)")


def _natural_key(value):
    """Natural alphanumeric ordering — spec §9.2 S1.

    Splits a label into runs of digits and non-digits; digit runs compare as numbers,
    the rest as text. That keeps ``2026`` and ``2026 9 months`` adjacent — they are two
    records of the same year, not a year and an outlier — while still ordering ``999``
    before ``2021``, which a plain string sort would not.
    """
    if value is None:
        return ((2, 0, ""),)                       # absences last

    parts = []
    for chunk in DIGITS.split(str(value).strip()):
        if not chunk:
            continue
        parts.append((0, int(chunk), "") if chunk.isdigit() else (1, 0, chunk.casefold()))
    return tuple(parts) or ((1, 0, ""),)


def _period_key(value, order: dict[str, int]):
    """Leading number, then declared suffix rank — spec §9.2 S13.

    Plain alphanumeric would order ``2025 9 months`` before ``2025 est``, because
    ``9`` sorts before ``e``. The declared order in ⟦PERIOD ORDER⟧ says otherwise:
    within one year, ``est`` precedes ``9 months`` precedes ``re-est``. A bare year
    sorts first; a suffix nobody declared sorts last, then alphanumerically among
    its kind, so the ordering stays total whatever a pack contains.
    """
    from .crosschecks import split_label

    number, suffix = split_label(value)
    if number is None:
        return None
    if not suffix:
        rank = 0
    else:
        rank = order.get(suffix.casefold(), max(order.values(), default=0) + 1)
    return ((0, number, ""), (0, rank, ""), *_natural_key(suffix))


def _sort_key(record: Record, fields, order=None):
    key = []
    for f in fields:
        value = record.values.get(f)
        period = _period_key(value, order) if order else None
        key.append(period if period is not None else _natural_key(value))
    return key


def _derived_figures(block: Block, spec: Step2Spec, actual_year: int | None):
    """Figures relating two records — resolved by leading number, then suffix."""
    from .crosschecks import match_record

    out = []
    for figure in spec.figures:
        if actual_year is None:
            out.append(Figure(figure.name, None,
                              "no 'Actual year' in ⟦GLOBAL⟧, so {N} cannot resolve",
                              figure.number_format, figure.note))
            continue

        num = match_record(block, figure.numerator, actual_year)
        den = match_record(block, figure.denominator, actual_year)
        missing = [p for p, r in ((figure.numerator, num), (figure.denominator, den))
                   if r is None]
        if missing:
            out.append(Figure(figure.name, None,
                              f"no record matching {' and '.join(missing)}",
                              figure.number_format, figure.note))
            continue

        a, b = num.values.get(figure.field), den.values.get(figure.field)
        if not isinstance(a, (int, float)) or not isinstance(b, (int, float)) or b == 0:
            out.append(Figure(figure.name, None,
                              f"{figure.field} missing or zero on one of the records",
                              figure.number_format, figure.note))
            continue

        value = a / b - 1 if figure.kind == "ratio_minus_1" else a / b
        detail = (f"{num.values[block.dataset.key_field]} {a:,.0f} ÷ "
                  f"{den.values[block.dataset.key_field]} {b:,.0f} − 1")
        out.append(Figure(figure.name, value, detail, figure.number_format, figure.note))
    return out


def _aggregate(block: Block, spec: AggregateSpec, blocks) -> AggregateTable:
    """Group and sum, spanning the declared year window — spec §9.2.1 S14."""
    from .crosschecks import split_label, years_in

    totals: dict[str, dict[str, float]] = {}
    for record in block.records:
        key = record.values.get(spec.group_by)
        if key is None:
            continue
        bucket = totals.setdefault(str(key), {m: 0.0 for m in spec.measures})
        for measure in spec.measures:
            value = record.values.get(measure)
            if isinstance(value, (int, float)):
                bucket[measure] += float(value)

    # A year with no record must still appear, showing 0.
    zero_filled = []
    window = None
    if blocks and spec.zero_fill_from:
        from .crosschecks import find_block
        pool = list(blocks.values()) if isinstance(blocks, dict) else list(blocks)
        window = find_block(pool, spec.zero_fill_from, block.section)
    if window is not None:
        present = {split_label(k)[0] for k in totals}
        for year in years_in(window):
            if year not in present:
                totals[str(year)] = {m: 0.0 for m in spec.measures}
                zero_filled.append(str(year))

    rows = sorted(totals.items(), key=lambda kv: _natural_key(kv[0]))
    return AggregateTable(spec.title, spec.group_by, spec.measures, rows, sorted(zero_filled))


def apply_step2(block: Block, spec: Step2Spec, nomenclature=None, blocks=None) -> Step2Result:
    missing = [f for f in spec.sort_by if f not in block.dataset.headers]
    if missing:
        raise ExtractionError(f"step 2 sorts by {missing}, which the dataset does not declare")

    actual_year = getattr(nomenclature, "actual_year", None)
    order = getattr(nomenclature, "period_order", None) or {}

    records = sorted(block.records, key=lambda r: _sort_key(r, spec.sort_by, order),
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

    sort_note = (f"sorted by {', '.join(spec.sort_by)} "
                 f"{'ascending' if spec.ascending else 'descending'}, natural alphanumeric")
    if order:
        ranked = " → ".join(s for s, _ in sorted(order.items(), key=lambda kv: kv[1]))
        sort_note += f", period order {ranked}"
    notes = [
        f"{sort_note} (value-preserving)",
        f"columns ordered as declared in sheet 00: "
        f"{' | '.join(block.dataset.headers)} (value-preserving)",
    ]
    for calc in spec.calculations:
        readable = FIELD.sub(lambda m: m.group(1), calc.expression).replace("/", " / ")
        notes.append(f"calculated {calc.name} = {readable} (value-adding)")

    aggregate = _aggregate(block, spec.aggregate, blocks) if spec.aggregate else None
    if aggregate is not None:
        notes.append(
            f"aggregated: {aggregate.title} — {spec.aggregate.group_by} × "
            f"{', '.join(spec.aggregate.measures)} (value-preserving in total)"
        )
        if aggregate.zero_filled:
            notes.append(
                f"years with no record shown as 0: {', '.join(aggregate.zero_filled)}"
            )

    result = Step2Result(
        block=block, spec=spec, records=records, computed=computed, notes=notes,
        figures=_derived_figures(block, spec, actual_year), aggregate=aggregate,
    )

    if aggregate is not None:
        for measure, total in aggregate.totals().items():
            if abs(total - result.totals().get(measure, 0.0)) > 1e-9:
                raise ExtractionError(
                    f"the aggregate total for {measure!r} is {total}, but the detail "
                    f"totals {result.totals().get(measure)} — grouping lost records"
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
