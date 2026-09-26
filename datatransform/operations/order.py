"""Ordering. Specification_v1.md §9.2 S1, S13.

Sorting is the operation everything downstream leans on without saying so: a running
share means nothing in an arbitrary order, and "the year before" means nothing at all. So
it is declared like any other step, and the operations that need it are declared after it.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from ..model import Record
from . import Context

DIGITS = re.compile(r"(\d+)")


def natural_key(value):
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


def period_key(value, order: dict[str, int]):
    """Leading number, then declared suffix rank — spec §9.2 S13.

    Plain alphanumeric would order ``2025 9 months`` before ``2025 est``, because
    ``9`` sorts before ``e``. The declared order in ⟦PERIOD ORDER⟧ says otherwise:
    within one year, ``est`` precedes ``9 months`` precedes ``re-est``. A bare year
    sorts first; a suffix nobody declared sorts last, then alphanumerically among
    its kind, so the ordering stays total whatever a pack contains.
    """
    from ..crosschecks import split_label

    number, suffix = split_label(value)
    if number is None:
        return None
    rank = 0 if not suffix else order.get(suffix.casefold(),
                                         max(order.values(), default=0) + 1)
    return ((0, number, ""), (0, rank, ""), *natural_key(suffix))


def sort_key(record: Record, fields, order=None, numeric=False):
    if numeric:
        # A bound is a number, so it orders as one. Absences sort first: a band open at
        # the bottom precedes every closed band.
        return [(record.values.get(f) is not None, record.values.get(f) or 0.0)
                for f in fields]
    key = []
    for f in fields:
        value = record.values.get(f)
        period = period_key(value, order) if order else None
        key.append(period if period is not None else natural_key(value))
    return key


@dataclass(frozen=True)
class SortBy:
    """Put the records in order — spec §9.2 S1, S13.

    ``numeric`` sorts on the value rather than the label, which is the whole reason band
    bounds are extracted as numbers: natural alphanumeric would place ``10 001 – 25 000``
    before ``5 001 – 10 000``.
    """

    fields: tuple[str, ...]
    ascending: bool = True
    numeric: bool = False

    def apply(self, ctx: Context) -> None:
        missing = [f for f in self.fields if f not in ctx.block.dataset.headers]
        if missing:
            from ..model import ExtractionError

            raise ExtractionError(
                f"step 2 sorts by {missing}, which the dataset does not declare"
            )

        order = ctx.period_order
        ctx.records = sorted(
            ctx.records,
            key=lambda r: sort_key(r, self.fields, order, self.numeric),
            reverse=not self.ascending,
        )

        note = (f"sorted by {', '.join(self.fields)} "
                f"{'ascending' if self.ascending else 'descending'}, "
                + ("numeric" if self.numeric else "natural alphanumeric"))
        if order and not self.numeric:
            ranked = " → ".join(s for s, _ in sorted(order.items(), key=lambda kv: kv[1]))
            note += f", period order {ranked}"
        ctx.notes.append(f"{note} (value-preserving)")

    def not_summable(self) -> set[str]:
        return set()
