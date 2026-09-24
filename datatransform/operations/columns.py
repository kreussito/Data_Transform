"""Columns step 2 adds. Specification_v1.md §9.2.1 S17, S21, §2.7.

Three ways a column can be worked out, and the difference between them is *what each one
looks at*:

============ ======================================= =====================================
``Calculate`` one record                             ``Incurred ÷ Premium``
``Change``    the record **before** it               a rate against last year's — §2.7
``Cumulative`` every record up to this one           where the book sits — S17
============ ======================================= =====================================

All three run after :class:`~datatransform.operations.order.SortBy`, and two of them are
meaningless without it. That dependency used to live in the order the code happened to
run; now it is in the declaration, where it can be read.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from ..constants import FMT_PERCENT
from ..nomenclature import norm
from . import Column, Context

FIELD = re.compile(r"\{([^}]+)\}")


@dataclass(frozen=True)
class Calculate:
    """A measure worked out from other fields of the same record — spec §9.2.1.

    Written into the sheet as a **live formula** over the columns beside it, so a reviewer
    can see the arithmetic rather than a number that claims to be its result.
    """

    name: str
    expression: str                       # fields in braces: "{Incurred Losses}/{Premium}"
    number_format: str = FMT_PERCENT
    guard_zero: str | None = None         # a field that must not be zero

    def apply(self, ctx: Context) -> None:
        values = []
        for record in ctx.records:
            guard = record.values.get(self.guard_zero) if self.guard_zero else None
            if self.guard_zero and guard in (None, 0):
                values.append(None)
                continue
            try:
                resolved = FIELD.sub(
                    lambda m: repr(float(record.values[m.group(1)])), self.expression)
                values.append(eval(resolved, {"__builtins__": {}}, {}))  # noqa: S307
            except (KeyError, TypeError, ValueError, ZeroDivisionError):
                values.append(None)

        ctx.derived.append(Column(self.name, values, self.number_format,
                                  formula=self.expression))
        readable = FIELD.sub(lambda m: m.group(1), self.expression).replace("/", " / ")
        ctx.notes.append(f"calculated {self.name} = {readable} (value-adding)")

    def not_summable(self) -> set[str]:
        return set()


@dataclass(frozen=True)
class Change:
    """A column derived from the same field on the **previous** record — spec §2.7, S21.

    Every other derivation in step 2 works inside one record or between two named ones.
    This one works along the *sequence*: a rate change is what the rate did since last
    year, so it needs the record before it and nothing else.

    Where the source already supplies ``name`` this does nothing — the underwriter's own
    figure is never replaced by one the tool worked out.
    """

    name: str
    field: str
    number_format: str = FMT_PERCENT
    key: str = "Year"                  # what must be consecutive for the change to be annual

    def apply(self, ctx: Context) -> None:
        if self.name in ctx.block.fields:
            ctx.notes.append(
                f"{self.name} as reported — {self.field} was not used to work it out")
            return
        if self.field not in ctx.block.fields:
            return                      # neither column: nothing to say, nothing to do

        values, gaps, previous = [], [], None
        for record in ctx.records:
            value = record.values.get(self.field)
            key = norm(record.values.get(self.key))
            if not isinstance(value, (int, float)) or previous is None:
                values.append(None)
            else:
                values.append(value / previous[1] - 1 if previous[1] else None)
                if _gap(previous[0], key) > 1:
                    gaps.append(f"{previous[0]}→{key}")
            record.values[self.name] = values[-1]
            if isinstance(value, (int, float)):
                previous = (key, value)

        # A declared field that step 2 worked out — so it belongs among the declared
        # columns, not appended after them. Its values live on the records.
        ctx.derive(self.name)
        ctx.block.display_formats.setdefault(self.name, self.number_format)
        ctx.notes.append(
            f"{self.name} worked out from {self.field}: each year against the one "
            f"before, as a **ratio** — so the unit of {self.field} does not matter "
            f"(%, ‰ or a bare number all give the same change), and the result is a "
            f"relative movement, never a difference in percentage points "
            f"(value-adding — the source gives the level, not the movement)"
        )
        if gaps:
            ctx.notes.append(
                f"{self.name} spans more than one year at {', '.join(gaps)} — the "
                f"{self.key.lower()}s are not consecutive there, so those figures are not "
                "annual changes and must not be read as such"
            )

    def not_summable(self) -> set[str]:
        """Five years' rates do not add up to a rate, nor five changes to a change."""
        return {self.name, self.field}


def _gap(a: str, b: str) -> int:
    """How many years apart two keys are; 1 where they cannot be read as numbers."""
    try:
        return int(float(b)) - int(float(a))
    except (TypeError, ValueError):
        return 1


@dataclass(frozen=True)
class Cumulative:
    """A running share of the column's total — spec §9.2.1 S17.

    What a profile exists to answer is "where does the book sit?", and that is a
    cumulative question: no single band answers it. Written as a live range so the reader
    sees both the running sum and the total it is divided by.
    """

    name: str
    field: str
    number_format: str = FMT_PERCENT

    def apply(self, ctx: Context) -> None:
        total = sum(r.values[self.field] for r in ctx.records
                    if isinstance(r.values.get(self.field), (int, float)))
        running, values = 0.0, []
        for record in ctx.records:
            value = record.values.get(self.field)
            if isinstance(value, (int, float)):
                running += float(value)
            values.append(running / total if total else None)

        ctx.derived.append(Column(self.name, values, self.number_format,
                                  running_of=self.field))
        ctx.notes.append(
            f"calculated {self.name} = running {self.field} ÷ total {self.field} "
            "(value-adding)"
        )

    def not_summable(self) -> set[str]:
        return set()
