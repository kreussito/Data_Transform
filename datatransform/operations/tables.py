"""Further step-2 tables and block-level figures. Spec §9.2.1 S12, S14, S15.

Both produce something that is *not* a column, and for the same reason: a figure relating
two records has no per-row meaning, and a grouping is a second table over the same data.
Writing either as a column would invite a reader to line it up against a record it does
not describe.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from ..constants import FMT_PERCENT
from ..model import Block, ExtractionError
from . import AggregateTable, Context, Figure
from .order import natural_key

FIELD = re.compile(r"\{([^}]+)\}")


@dataclass(frozen=True)
class DerivedFigure:
    """A figure relating two *records* rather than describing one — spec §9.2.1 S12.

    Written beneath the data, not as a column, because it has no per-row meaning. The two
    records are named by pattern — ``{N} re-est`` against ``{N} est`` — and resolved
    against ⟦GLOBAL⟧'s actual year.
    """

    name: str
    numerator: str
    denominator: str
    field: str
    kind: str = "ratio_minus_1"
    number_format: str = FMT_PERCENT
    note: str = ""

    def apply(self, ctx: Context) -> None:
        ctx.figures.extend(_figures(ctx.block, [self], ctx.actual_year))

    def not_summable(self) -> set[str]:
        return set()


def _figures(block: Block, specs, actual_year: int | None):
    """Figures relating two records — resolved by leading number, then suffix."""
    from ..crosschecks import match_record

    out = []
    for figure in specs:
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


DERIVED_GROUP = re.compile(r"^year\((.+)\)$", re.I)


def _group_key(record, group_by: str):
    """The value a record groups under — a field, or ``year(<date field>)``.

    An event may span underwriting years, so grouping by the *occurrence* year means
    reading it off the date rather than the declared year — spec §9.2.1 S15.
    """
    derived = DERIVED_GROUP.match(group_by.strip())
    if not derived:
        return record.values.get(group_by)
    value = record.values.get(derived.group(1).strip())
    return str(value.year) if hasattr(value, "year") else None


@dataclass(frozen=True)
class Aggregate:
    """A second step-2 table: group by a key, sum measures — spec §9.2.1 S14, S15.

    ``zero_fill_from`` names a dataset whose year window the table should span, so a year
    with no records shows 0 rather than being absent. A missing year reads as "no data";
    a zero reads as "nothing happened", and only one of those is true.

    ``group_by_attribute`` names an attribute that may redirect the grouping — which date
    defines the occurrence year is the cedent's convention, not the tool's.
    """

    title: str
    group_by: str
    measures: tuple[str, ...]
    zero_fill_from: str | None = None
    group_by_attribute: str | None = None
    note: str = ""

    def apply(self, ctx: Context) -> None:
        table = _aggregate(ctx.block, self, ctx.blocks)
        ctx.aggregates.append(table)
        ctx.notes.append(
            f"aggregated: {table.title} — {table.group_by} × "
            f"{', '.join(table.measures)} (value-preserving in total)")
        if table.zero_filled:
            ctx.notes.append(
                f"{table.title}: groups with no record shown as 0: "
                f"{', '.join(table.zero_filled)}")

    def not_summable(self) -> set[str]:
        return set()


def _resolve_group_by(block: Block, spec) -> str:
    """The grouping actually used, after any attribute redirects it — spec §9.2.1 S15.

    ``year(Event Date)`` is the tool's default reading of "occurrence year". A cedent
    who defines it by when the event *ended* says so in column A, and the table follows
    without the code being touched.
    """
    if not spec.group_by_attribute:
        return spec.group_by
    attribute = block.attributes.get(spec.group_by_attribute)
    if attribute is None:
        return spec.group_by

    field = str(attribute.value).strip()
    if field not in block.fields:
        raise ExtractionError(
            f"{block.sheet_name!r} block {block.index}: "
            f"{spec.group_by_attribute} = {field!r}, which the block does not extract "
            f"(it has {', '.join(block.fields)})"
        )
    derived = DERIVED_GROUP.match(spec.group_by.strip())
    return f"year({field})" if derived else field


def _aggregate(block: Block, spec, blocks) -> AggregateTable:
    """Group and sum, spanning the declared year window — spec §9.2.1 S14."""
    from ..crosschecks import split_label, years_in

    group_by = _resolve_group_by(block, spec)
    totals: dict[str, dict[str, float]] = {}
    for record in block.records:
        key = _group_key(record, group_by)
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
        from ..crosschecks import find_block
        pool = list(blocks.values()) if isinstance(blocks, dict) else list(blocks)
        window = find_block(pool, spec.zero_fill_from, block.section)
    if window is not None:
        present = {split_label(k)[0] for k in totals}
        for year in years_in(window):
            if year not in present:
                totals[str(year)] = {m: 0.0 for m in spec.measures}
                zero_filled.append(str(year))

    rows = sorted(totals.items(), key=lambda kv: natural_key(kv[0]))
    return AggregateTable(spec.title, group_by, spec.measures, rows,
                          sorted(zero_filled), spec.note)


