"""The split bridge. Specification_v1.md §2.6, §9.2.1 S20.

A cat model wants sums insured on a fixed grid. A cedent reports on whatever grid it keeps
its book: one figure per zone, a cover split, an occupancy split, or a Projects/Renewables
segmentation that is not one of the target axes at all. The arithmetic that gets from one
to the other lives in :mod:`datatransform.bridge`; what lives here is the *operation* —
choosing the level, applying the bridge, and saying in the sheet what was done.

Declared **before** :class:`~datatransform.operations.reconcile.Identity`, because the
identity has nothing to reconcile until the buckets exist. That used to be a comment
above a line of code; now it is the order in the pipeline.
"""

from __future__ import annotations

from dataclasses import dataclass
from itertools import product as _product

from ..bridge import LevelFinding, build_bridge, fit_margins, view_threshold
from ..constants import (
    F_TOTAL,
    LEVEL_BOTH,
    LEVEL_COVER,
    LEVEL_OCCUPANCY,
    LEVEL_REPORTED,
    LEVEL_SEGMENT,
    LEVEL_TOTAL,
)
from ..model import Block, ExtractionError, Record
from . import Context


def _reported_level(block: Block, occupancy, covers, segments, total: str):
    """What the source actually reports, best first — spec §2.6.

    The order is the order of information: both margins beat one, one beats a
    segmentation that has to be translated, and that beats a bare total.
    """
    have = set(block.fields)
    occ = [o for o in occupancy if o in have]
    cov = [c for c in covers if c in have]
    seg = [s for s in segments if s in have]

    if occ and cov:
        return LEVEL_BOTH, (tuple(occ), tuple(cov))
    if occ:
        return LEVEL_OCCUPANCY, tuple(occ)
    if cov:
        return LEVEL_COVER, tuple(cov)
    if seg:
        return LEVEL_SEGMENT, tuple(seg)
    if total in have:
        return LEVEL_TOTAL, (total,)
    return "", ()


def _unused_levels(block: Block, targets, total: str, kind: str, labels) -> list[str]:
    """Reported columns the chosen level does not consume."""
    used = set(labels[0] + labels[1]) if kind == LEVEL_BOTH else set(labels)
    return [f for f in block.numeric_fields
            if f not in used and f != total and f not in set(targets)]


def _level_of(field_name: str, occupancy, covers, segments) -> str:
    for kind, members in ((LEVEL_OCCUPANCY, occupancy), (LEVEL_COVER, covers),
                          (LEVEL_SEGMENT, segments)):
        if field_name in members:
            return kind
    return ""


def _reconcile_levels(block: Block, records, bridge, occupancy, covers, segments,
                      targets, total: str, kind: str, labels, nomenclature=None):
    """Two descriptions of the same book — compared, not silently resolved. Spec §2.6.

    A cedent may send the occupancy split *and* a Projects/Renewables split. Each implies
    an occupancy mix, and nothing in the numbers says which one the cedent stands behind.
    Choosing one quietly would lose the other without a trace; refusing outright would
    throw away a perfectly good submission over a question a person can answer in a
    sentence. So the block is built from the richer level, the other is bridged too, and
    the two are shown side by side with the question put to the underwriter in writing.
    """
    unused = _unused_levels(block, targets, total, kind, labels)
    if not unused:
        return None

    kinds = {_level_of(f, occupancy, covers, segments) for f in unused}
    if not kinds or "" in kinds or len(kinds) > 1:
        raise ExtractionError(
            f"{block.sheet_name!r} block {block.index}: the block reports "
            f"{', '.join(unused)}, which belongs to no level sheet 00 declares — so it "
            "can be neither used nor compared, and dropping a column that carries money "
            "is not something this step will do in silence"
        )

    secondary = kinds.pop()
    finding = LevelFinding(primary=kind, secondary=secondary,
                           labels=tuple(occupancy),
                           threshold=view_threshold(nomenclature))

    used = labels[0] + labels[1] if kind == LEVEL_BOTH else tuple(labels)
    mine = {o: 0.0 for o in occupancy}
    theirs = {o: 0.0 for o in occupancy}

    for record in records:
        if kind == LEVEL_REPORTED:
            # The grid itself is the primary view; collapse it onto the occupancy axis.
            for name in targets:
                head = next((o for o in occupancy
                             if name == o or name.startswith(f"{o} ")), None)
                if head:
                    mine[head] += _amount(record, name)
        else:
            for label in used:
                level = _level_of(label, occupancy, covers, segments) or LEVEL_TOTAL
                for (o, _), share in bridge.distribute(level, label).items():
                    mine[o] += _amount(record, label) * share
        for label in unused:
            for (o, _), share in bridge.distribute(secondary, label).items():
                theirs[o] += _amount(record, label) * share

    finding.rows = [(o, mine[o], theirs[o]) for o in occupancy]
    return finding


@dataclass(frozen=True)
class Split:
    """Bridge what arrived onto the dataset's target cells — spec §2.6, S20.

    Two steps, and keeping them apart is what makes every case one case: **expand** what
    the source reports onto the section's occupancy × cover grid, then **project** that
    grid onto the axes the dataset actually declares. Earthquake declares both, so
    nothing is projected away; windstorm declares occupancy alone, so the cover axis is
    summed out — which still uses the cover information rather than discarding it.

    A split redistributes and never creates: the record total is unchanged.
    """

    total: str = F_TOTAL

    def apply(self, ctx: Context) -> None:
        block, nomenclature = ctx.block, ctx.nomenclature
        if nomenclature is None:
            return

        axes = nomenclature.axes_for(block.dataset.key)
        targets = nomenclature.buckets_for(block.dataset.key)
        if not axes or not targets:
            return
        segments = tuple(nomenclature.segment_categories())
        occupancy = tuple(axes[0].categories)
        target_covers = tuple(axes[1].categories) if len(axes) > 1 else ()
        covers = target_covers or nomenclature.cover_categories()

        if all(t in block.fields for t in targets):
            # The finished grid arrived. Any *other* level column describes the same book
            # a second way, and choosing between them silently is not the tool's to do.
            if _unused_levels(block, targets, self.total, LEVEL_REPORTED, targets):
                bridge = build_bridge(nomenclature, ctx.blocks or [], block.section,
                                      occupancy, covers)
                if bridge is not None:
                    ctx.level_finding = _reconcile_levels(
                        block, ctx.records, bridge, occupancy, bridge.covers, segments,
                        targets, self.total, LEVEL_REPORTED, targets, nomenclature)
            return

        bridge = build_bridge(nomenclature, ctx.blocks or [], block.section,
                              occupancy, covers)
        if bridge is None:
            self._from_declared_ratios(ctx, targets, axes)
            return

        kind, labels = _reported_level(block, occupancy, bridge.covers, segments,
                                       self.total)
        if not kind:
            raise ExtractionError(
                f"{block.sheet_name!r} block {block.index}: the source reports none of "
                f"the {len(targets)} target cell(s), neither margin, no segmentation and "
                f"no {self.total} — there is no level to build from, and the tool will "
                "not invent one")

        out = []
        projection = _project(targets, axes, occupancy, bridge.covers)
        for record in ctx.records:
            values = dict(record.values)
            grid = _expand(record, kind, labels, bridge, occupancy)
            for name, cell in projection.items():
                values[name] = sum(grid.get(k, 0.0) for k in cell)
            out.append(Record(record.source_ref, values, record.confidence))

        created = tuple(t for t in targets if t not in block.fields)
        ctx.records = out
        ctx.derive(*created)
        ctx.assumed_fields += created
        ctx.flags["split_fields"] = created
        ctx.flags["parts_from_total"] = kind == LEVEL_TOTAL
        ctx.notes.extend(_split_notes(kind, labels, bridge, targets, axes))
        ctx.level_finding = _reconcile_levels(
            block, ctx.records, bridge, occupancy, bridge.covers, segments, targets,
            self.total, kind, labels, nomenclature)

    def _from_declared_ratios(self, ctx: Context, targets, axes) -> None:
        """No 08 for this section — fall back to ratios typed into ⟦SPLITS⟧ — spec §2.6."""
        block, nomenclature = ctx.block, ctx.nomenclature
        if self.total not in block.fields:
            raise ExtractionError(
                f"{block.sheet_name!r} block {block.index}: section {block.section!r} "
                f"carries no 08 split table and the block reports no {self.total}, so "
                "there is nothing to build the target cells from")

        shares, sources = {}, []
        for name in targets:
            factor, origin = 1.0, []
            for axis, category in zip(axes, _cells_of(name, axes)):
                declared = {r.category: r for r in
                            nomenclature.splits_for(block.dataset.key, axis.name)}
                if category not in declared:
                    raise ExtractionError(
                        f"{block.sheet_name!r} block {block.index}: no 08 table and sheet "
                        f"00 ⟦SPLITS⟧ declares no share for {category!r} on axis "
                        f"{axis.name!r} — the tool will not invent one")
                factor *= declared[category].share
                if declared[category].source:
                    origin.append(declared[category].source)
            shares[name] = factor
            sources.extend(origin)

        ctx.records = [
            Record(r.source_ref,
                   {**r.values, **{n: _amount(r, self.total) * s
                                   for n, s in shares.items()}},
                   r.confidence)
            for r in ctx.records
        ]
        created = tuple(t for t in targets if t not in block.fields)
        ctx.derive(*created)
        ctx.assumed_fields += created
        ctx.flags["split_fields"] = created
        ctx.flags["parts_from_total"] = True

        origin = "; ".join(dict.fromkeys(sources))
        ctx.notes.append(
            f"no 08 split table for section {block.section!r}; {len(targets)} target "
            f"cell(s) built from ratios declared in sheet 00 ⟦SPLITS⟧"
            + (f" [{origin}]" if origin else "")
            + " (value-adding — an assumption, not a reading)")
        ctx.notes.append(
            "the split redistributes and leaves every zone total unchanged")

    def not_summable(self) -> set[str]:
        return set()


def _expand(record, kind: str, labels, bridge, occupancy) -> dict:
    """One record's amounts, spread over the section's occupancy × cover grid."""
    if kind == LEVEL_BOTH:
        occ_labels, cover_labels = labels
        return fit_margins(
            bridge.joint(), occupancy, bridge.covers,
            {o: _amount(record, o) for o in occupancy},
            {c: _amount(record, c) for c in bridge.covers},
        ) if all(o in occ_labels for o in occupancy) else fit_margins(
            bridge.joint(), occ_labels, cover_labels,
            {o: _amount(record, o) for o in occ_labels},
            {c: _amount(record, c) for c in cover_labels},
        )

    grid: dict = {}
    for label in labels:
        amount = _amount(record, label)
        source = LEVEL_TOTAL if kind == LEVEL_TOTAL else kind
        for cell, share in bridge.distribute(source, label).items():
            grid[cell] = grid.get(cell, 0.0) + amount * share
    return grid


def _amount(record, field_name: str) -> float:
    value = record.values.get(field_name)
    return float(value) if isinstance(value, (int, float)) else 0.0


def _project(targets, axes, occupancy, covers) -> dict:
    """Target cell name → the grid keys it collects. Windstorm sums the cover axis out."""
    out = {}
    if len(axes) > 1:
        for name in targets:
            for o in occupancy:
                for c in covers:
                    if name == f"{o} {c}":
                        out[name] = [(o, c)]
    else:
        for name in targets:
            out[name] = [(name, c) for c in covers]
    return out


def _split_notes(kind, labels, bridge, targets, axes) -> list[str]:
    told = {
        LEVEL_BOTH: "both margins reported per zone, fitted to the grid from {src} so that "
                "neither margin moves",
        LEVEL_OCCUPANCY: "occupancy reported per zone; the cover mix of each comes from {src}",
        LEVEL_COVER: "cover reported per zone; the occupancy mix of each comes from {src} — "
                 "which is why residential BI stays near nil rather than taking a flat share",
        LEVEL_SEGMENT: "reported as {shown}, translated onto the target cells through {src} "
                   "and the declared occupancy convention",
        LEVEL_TOTAL: "one figure per zone, spread over the target cells by the joint "
                 "distribution of {src}",
    }[kind]
    shown = ", ".join(labels if kind != LEVEL_BOTH else labels[0] + labels[1])
    notes = [
        f"{len(targets)} target cell(s) built: "
        + told.format(src=bridge.source, shown=shown)
        + " (value-adding — an assumption about the zone, not about the book)"
    ]
    if kind == LEVEL_SEGMENT and bridge.conventions:
        notes.append(
            "occupancy of " + " and ".join(bridge.conventions)
            + " is a declared convention in sheet 00 ⟦SPLITS⟧, not a reading — sheet 08 "
              "cannot supply that edge"
        )
    if len(axes) == 1:
        notes.append(
            f"the cover axis is summed out: this dataset declares {axes[0].name} only, "
            "but the cover figures still informed how the occupancy was worked out"
        )
    notes.append(
        "the split redistributes and leaves every zone total unchanged"
    )
    return notes


def _cells_of(name: str, axes) -> list[str]:
    """Split a target cell name back into one category per axis."""
    if len(axes) == 1:
        return [name]
    for first in axes[0].categories:
        if name.startswith(f"{first} ") and name[len(first) + 1:] in axes[1].categories:
            return [first, name[len(first) + 1:]]
    return [name]


