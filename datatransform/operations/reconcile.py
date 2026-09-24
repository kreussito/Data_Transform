"""Reconciling and completing a record set. Specification_v1.md §9.2.1 S18, S19.

Two operations that change *what is in the table* rather than what is in a record:
:class:`Identity` reconciles a total against its parts, :class:`Complete` adds the
catalogue members the source never mentioned.

Both are value-preserving in the strict sense — a derived total is arithmetic over
figures already there, and adding zeros cannot move a sum — which is why S6's control
check proves them on every run rather than trusting them.
"""

from __future__ import annotations

from dataclasses import dataclass

from ..constants import FINDING_EXAMPLES, IDENTITY_TOLERANCE
from ..model import ExtractionError, Record
from ..nomenclature import norm
from . import Context


@dataclass(frozen=True)
class Identity:
    """A declared field that equals the sum of others — spec §9.2.1 S19.

    Derived where the source omits it, **checked** where the source supplies it, and left
    alone where the parts are missing — the same three shapes as the band bounds. A total
    that disagrees with its parts is reported and left exactly as the cedent sent it: the
    tool does not know which of the figures is the wrong one, and a quietly corrected
    total is worse than a visible contradiction.

    ``parts`` empty means *the dataset's buckets*, taken from ⟦AXES⟧ — so a dataset whose
    axes change does not need its identity restated.
    """

    total: str
    parts: tuple[str, ...] = ()

    def apply(self, ctx: Context) -> None:
        block = ctx.block
        declared = self.parts or (
            ctx.nomenclature.buckets_for(block.dataset.key) if ctx.nomenclature else ())
        if not declared:
            return

        known = (set(block.fields) | set(ctx.records[0].values)) if ctx.records \
            else set(block.fields)
        parts = [p for p in declared if p in known]
        if not parts:
            ctx.notes.append(
                f"{self.total} stands alone: the source declares none of "
                f"{declared[0]} … {declared[-1]}, so the split cannot be reconciled")
            return

        if ctx.flags.get("parts_from_total"):
            ctx.notes.append(
                f"{self.total} was the source of the {len(parts)} part(s), so they sum "
                "back to it by construction — nothing here is a check")
            return

        if self.total in block.fields:
            off = []
            for record in ctx.records:
                total = record.values.get(self.total)
                summed = sum(record.values[p] for p in parts
                             if isinstance(record.values.get(p), (int, float)))
                if isinstance(total, (int, float)) \
                        and abs(total - summed) > IDENTITY_TOLERANCE:
                    off.append(f"{record.values.get(self.total, '')}"
                               f"{record.source_ref}: {total:,.0f} vs {summed:,.0f}")
            ctx.notes.append(
                f"{self.total} checked against {len(parts)} part(s): "
                + ("all agree" if not off
                   else "DISAGREES — " + "; ".join(off[:FINDING_EXAMPLES])))
            return

        ctx.records = [
            Record(r.source_ref,
                   {**r.values,
                    self.total: sum(r.values[p] for p in parts
                                    if isinstance(r.values.get(p), (int, float)))},
                   r.confidence)
            for r in ctx.records
        ]
        ctx.derive(self.total)
        ctx.notes.append(
            f"{self.total} derived as the sum of {len(parts)} declared part(s) — the "
            "source does not supply it (value-adding, but arithmetic only)")

    def not_summable(self) -> set[str]:
        return set()


@dataclass(frozen=True)
class Complete:
    """Every member of a declared catalogue must appear — spec §9.2.1 S18.

    A cat aggregate lists only the zones the cedent has exposure in. The zones with none
    are exactly the ones worth seeing: an absent row reads as *no data*, a zero reads as
    *nothing there*, and only the second is a statement about the portfolio.

    ``catalogue_attribute`` names the attribute that selects which list applies —
    ``Zone scheme = Mexico EQ`` and ``Mexico Wind`` are different zonings of the same
    country and must not be mixed.
    """

    key: str
    catalogue_attribute: str

    def apply(self, ctx: Context) -> None:
        block = ctx.block
        if ctx.nomenclature is None or self.key not in block.fields:
            return
        declared = block.attributes.get(self.catalogue_attribute)
        if declared is None:
            return

        catalogue = ctx.nomenclature.zones_for(str(declared.value))
        if not catalogue:
            raise ExtractionError(
                f"{block.sheet_name!r} block {block.index}: "
                f"{self.catalogue_attribute} = {declared.value!r}, which sheet 00 "
                "⟦ZONES⟧ does not list — the members with no exposure cannot be shown "
                "without it")

        extra = ctx.flags.get("split_fields", ())
        present = {norm(r.values.get(self.key)) for r in ctx.records}
        filled = []
        for member in catalogue:
            if norm(member) in present:
                continue
            values = {self.key: member}
            values.update({m: 0.0 for m in (*block.measure_fields, *extra)})
            ctx.records.append(Record("—", values))
            filled.append(member)

        if filled:
            ctx.notes.append(
                f"{len(filled)} declared {self.key.lower()}(s) with no entry in the "
                f"source, shown as 0: {', '.join(filled)} (value-preserving)")

    def not_summable(self) -> set[str]:
        return set()
