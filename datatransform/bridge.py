"""From what the cedent reported onto the nine cells. Specification_v1.md §2.6.

A cat model wants sums insured on a fixed grid — occupancy × cover for earthquake,
occupancy alone for windstorm. A cedent reports on whatever grid it keeps its book: one
figure per zone, a cover split, an occupancy split, or — for an engineering book — a
segmentation into Projects and Renewables that is not one of the target axes at all.

So the operation is not "multiply the missing axis on". It is a **bridge**::

    B(z, t) = Σₐ  S(z, a) · M(a → t)          with   Σₜ M(a → t) = 1

``S`` is what arrived, ``M`` sends each reported category to a distribution over the
target cells. Every case is that one formula:

===================== ==================================================================
reported occupancy    ``M(Res → Res·j) = p(j|Res)``, zero outside the row
reported cover        ``M(Building → i·Building) = p(i|Building)``, zero outside the column
reported segment      ``M(Renewables → i·j) = q(i|Renewables) · p(j|Renewables)``
nothing but a total   ``M(Total → i·j) = p(i,j)``
===================== ==================================================================

The zeros are why a reported figure survives untouched: the row sums to exactly what
arrived. The segment axis has no such zeros, and that is precisely what distinguishes it
— it is a translation between two descriptions of the same book, not a refinement of one.

**Where M comes from.** The joint ``p(i,j)`` is read off sheet ``08``, the cedent's own
split table, per section. Shares taken from amounts cannot fail to close, and because
only ratios are used, ``08``'s scale and currency are irrelevant — only its internal
consistency matters.

The one edge no cedent supplies is ``q(i|segment)``: nobody cross-tabulates Renewables
against residential/commercial/industrial. That is a declared convention in ⟦SPLITS⟧ —
Renewables is industrial, a Project is half commercial and half industrial — and it is
marked as one wherever it is applied.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .model import ExtractionError
from .nomenclature import norm

TOTAL = "Total"
NO_COVER = ""          # 08 carrying no cover breakdown: one implicit cover category

IPF_ROUNDS = 60
IPF_TOLERANCE = 1e-9


@dataclass
class Bridge:
    """The joint distribution of one section's book, and the maps derived from it."""

    section: str
    occupancy: tuple[str, ...]
    covers: tuple[str, ...]
    amounts: dict[tuple[str, str], float]          # (occupancy, cover) → sum insured
    segments: dict[str, dict[str, float]] = field(default_factory=dict)
    segment_occupancy: dict[str, dict[str, float]] = field(default_factory=dict)
    source: str = ""
    conventions: list[str] = field(default_factory=list)

    # ── the distributions ────────────────────────────────────────────────────
    @property
    def total(self) -> float:
        return sum(self.amounts.values())

    def joint(self) -> dict[tuple[str, str], float]:
        """p(i,j) — used where nothing but a zone total arrived."""
        total = self.total
        return {k: v / total for k, v in self.amounts.items()} if total else {}

    def cover_given_occupancy(self, occupancy: str) -> dict[str, float]:
        """p(j|i) — the cover mix of one occupancy."""
        row = {c: self.amounts.get((occupancy, c), 0.0) for c in self.covers}
        return _normalise(row, f"cover given {occupancy!r}", self)

    def occupancy_given_cover(self, cover: str) -> dict[str, float]:
        """p(i|j) — the occupancy mix of one cover. Residential BI is nearly nil, and
        this is the map that knows it; a single occupancy vector applied to all three
        covers would not."""
        column = {o: self.amounts.get((o, cover), 0.0) for o in self.occupancy}
        return _normalise(column, f"occupancy given {cover!r}", self)

    def cover_given_segment(self, segment: str) -> dict[str, float]:
        """p(j|a) — read off 08's own row for that segment, where it has one."""
        row = self.segments.get(segment)
        if not row:
            return _normalise({c: sum(self.amounts.get((o, c), 0.0)
                                      for o in self.occupancy) for c in self.covers},
                              "cover overall", self)
        return _normalise(dict(row), f"cover given {segment!r}", self)

    def occupancy_given_segment(self, segment: str) -> dict[str, float]:
        """q(i|a) — the declared convention; 08 cannot supply this edge."""
        declared = self.segment_occupancy.get(segment)
        if not declared:
            raise ExtractionError(
                f"section {self.section!r}: nothing maps {segment!r} onto "
                f"{', '.join(self.occupancy)}. Sheet 08 cannot supply it — no cedent "
                "cross-tabulates that way — so it must be declared in sheet 00 ⟦SPLITS⟧"
            )
        return dict(declared)

    # ── the bridge itself ────────────────────────────────────────────────────
    def distribute(self, kind: str, label: str) -> dict[tuple[str, str], float]:
        """M(a → ·) for one reported category. The four rows of the table above."""
        if kind == "occupancy":
            return {(label, c): s for c, s in self.cover_given_occupancy(label).items()}
        if kind == "cover":
            return {(o, label): s for o, s in self.occupancy_given_cover(label).items()}
        if kind == "segment":
            occ = self.occupancy_given_segment(label)
            cov = self.cover_given_segment(label)
            return {(o, c): a * b for o, a in occ.items() for c, b in cov.items()}
        if kind == "total":
            return self.joint()
        raise ValueError(f"unknown reported level {kind!r}")


def _normalise(weights: dict[str, float], what: str, bridge: Bridge) -> dict[str, float]:
    total = sum(weights.values())
    if total <= 0:
        raise ExtractionError(
            f"section {bridge.section!r}: sheet 08 carries nothing under {what}, so the "
            "share cannot be worked out — and the tool will not spread it evenly instead"
        )
    return {k: v / total for k, v in weights.items()}


# ────────────────────────────────────────────────── building it from sheet 08
CATEGORY_FIELD = "Category"


def build_bridge(nomenclature, blocks, section: str, occupancy, covers) -> Bridge | None:
    """Read one section's split table. ``None`` where the section carries no 08."""
    from .crosschecks import find_block

    pool = list(blocks.values()) if isinstance(blocks, dict) else list(blocks)
    table = find_block(pool, "08", section)
    if table is None:
        return None

    known_occupancy = {norm(o).casefold(): o for o in occupancy}
    cover_fields = [c for c in covers if c in table.fields]
    if not cover_fields and TOTAL not in table.fields:
        raise ExtractionError(
            f"{table.sheet_name!r}: the split table declares neither the cover columns "
            f"({', '.join(covers)}) nor {TOTAL}, so it carries no amounts to take a "
            "share of"
        )

    amounts: dict[tuple[str, str], float] = {}
    segments: dict[str, dict[str, float]] = {}
    listed_covers = tuple(cover_fields) if cover_fields else (NO_COVER,)

    for record in table.records:
        label = norm(record.values.get(CATEGORY_FIELD))
        row = ({c: _number(record.values.get(c)) for c in cover_fields} if cover_fields
               else {NO_COVER: _number(record.values.get(TOTAL))})
        match = known_occupancy.get(label.casefold())
        if match is not None:
            for cover, value in row.items():
                amounts[(match, cover)] = amounts.get((match, cover), 0.0) + value
        else:
            segments[label] = row                 # Projects, Renewables — a third axis

    bridge = Bridge(
        section=section, occupancy=tuple(occupancy), covers=listed_covers,
        amounts=amounts, segments=segments,
        source=f"{table.sheet_name} block {table.index}",
    )
    _attach_conventions(bridge, nomenclature, table.dataset.key)

    if segments and not amounts:
        # An engineering book: 08 is keyed by Projects/Renewables, so the occupancy grid
        # has to be built through the declared convention before anything else can use it.
        _fold_segments(bridge)
    if not bridge.amounts:
        raise ExtractionError(
            f"{table.sheet_name!r}: no row matches {', '.join(occupancy)} and none could "
            "be mapped there, so the section has no occupancy split at all"
        )
    return bridge


def _attach_conventions(bridge: Bridge, nomenclature, dataset_key: str) -> None:
    """⟦SPLITS⟧ rows whose ``From`` names something outside the occupancy axis."""
    targets = set(bridge.occupancy)
    for rule in getattr(nomenclature, "splits", None) or []:
        if not rule.source_category or rule.category not in targets:
            continue
        if not (rule.applies_to(dataset_key) or rule.dataset == "*"):
            continue
        bridge.segment_occupancy.setdefault(rule.source_category, {})[rule.category] = \
            rule.share
        if rule.source:
            bridge.conventions.append(f"{rule.source_category} [{rule.source}]")
    bridge.conventions = list(dict.fromkeys(bridge.conventions))


def _fold_segments(bridge: Bridge) -> None:
    """Turn a Projects/Renewables table into the occupancy grid it stands for."""
    for segment, row in bridge.segments.items():
        occ = bridge.occupancy_given_segment(segment)
        for cover, value in row.items():
            for name, share in occ.items():
                key = (name, cover)
                bridge.amounts[key] = bridge.amounts.get(key, 0.0) + value * share


def _number(value) -> float:
    return float(value) if isinstance(value, (int, float)) else 0.0


# ─────────────────────────────────────────── two reported margins: fit both
def fit_margins(seed, rows, columns, row_totals, column_totals):
    """Fit a grid to **both** reported margins — spec §2.6.

    Where a zone reports occupancy *and* cover as two vectors, conditioning on one of
    them preserves that one and lets the other drift. Both were reported, so neither may
    move: the seed from 08 is scaled alternately to the row and column totals until it
    satisfies both (RAS/iterative proportional fitting).

    The seed decides only *how* the two margins interact — the margins themselves come
    out exactly as sent.
    """
    grid = {(r, c): max(seed.get((r, c), 0.0), 0.0) for r in rows for c in columns}
    if sum(grid.values()) <= 0:
        grid = {k: 1.0 for k in grid}

    for _ in range(IPF_ROUNDS):
        moved = 0.0
        for r in rows:
            current = sum(grid[(r, c)] for c in columns)
            if current > 0:
                factor = row_totals.get(r, 0.0) / current
                moved = max(moved, abs(factor - 1))
                for c in columns:
                    grid[(r, c)] *= factor
        for c in columns:
            current = sum(grid[(r, c)] for r in rows)
            if current > 0:
                factor = column_totals.get(c, 0.0) / current
                moved = max(moved, abs(factor - 1))
                for r in rows:
                    grid[(r, c)] *= factor
        if moved < IPF_TOLERANCE:
            break
    return grid
