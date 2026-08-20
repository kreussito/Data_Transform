"""Step 1 and step 2. Specification_v1.md §9.2, §10."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .bridge import build_bridge, fit_margins
from .model import Block, Confidence, ExtractionError, Hypothesis, Record
from .nomenclature import norm
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
    note: str = ""

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
    aggregates: list[AggregateTable] = field(default_factory=list)
    derived_fields: tuple[str, ...] = ()     # declared fields step 2 read off a label
    assumed_fields: tuple[str, ...] = ()     # fields resting on a declared ratio — §2.6

    @property
    def confidence(self) -> Confidence:
        """Step 1's confidence, but never better than the ratios it was expanded with.

        A block extracted cleanly is ``Confirmed``. If three of its five columns exist
        only because a declared split was multiplied onto a reported figure, the block as
        a whole is no longer a reading, and saying so is the entire point of carrying
        confidence at all — spec §5.3.
        """
        if not self.assumed_fields:
            return self.block.confidence
        return Confidence.worst((self.block.confidence, Confidence.ASSUMED))

    @property
    def declared_columns(self) -> tuple[str, ...]:
        """Exactly the order sheet 00 declares, less any field neither present nor derived."""
        available = set(self.block.fields) | set(self.derived_fields)
        return tuple(h for h in self.block.dataset.headers if h in available)

    @property
    def columns(self) -> tuple[str, ...]:
        """Declared columns, then derived ones — spec §9.2 S3."""
        return (self.declared_columns
                + tuple(c.name for c in self.spec.calculations)
                + tuple(c.name for c in self.spec.cumulative))

    def totals(self) -> dict[str, float]:
        out = {}
        for m in self.block.measure_fields:
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


def _sort_key(record: Record, fields, order=None, numeric=False):
    if numeric:
        # A bound is a number, so it orders as one. Absences sort first: a band open at
        # the bottom precedes every closed band.
        return [(record.values.get(f) is not None, record.values.get(f) or 0.0)
                for f in fields]
    key = []
    for f in fields:
        value = record.values.get(f)
        period = _period_key(value, order) if order else None
        key.append(period if period is not None else _natural_key(value))
    return key


def _span(record, bounds) -> str:
    """A band's name when the source gave no label — its bounds are its identity."""
    lower, upper = record.values.get(bounds.lower), record.values.get(bounds.upper)
    return (f"{'open' if lower is None else format(lower, ',.0f')} – "
            f"{'open' if upper is None else format(upper, ',.0f')}")


def _derive_bounds(block: Block, spec: Step2Spec):
    """Make sure step 2 always has both bounds, however the source supplied them.

    Three shapes, one outcome — spec §2.4:

    ==================================== ==================================================
    two numeric columns                  extracted; nothing is worked out
    a label only (``1 - 1,000,000``)     the bounds are read off it and reported
    two columns and no label             extracted; the band is named by its own bounds
    ==================================== ==================================================

    Returns ``(records, derived field names, continuity findings)``. Where anything is
    derived the records are **copies**: step 1 holds what the sheet says and must not
    gain a column step 2 worked out (S11).
    """
    from .bands import continuity, parse_band

    bounds = spec.bounds
    if bounds is None:
        return list(block.records), (), None

    has_label = bounds.label in block.fields
    missing = [f for f in (bounds.lower, bounds.upper) if f not in block.fields]

    if missing and not has_label:
        raise ExtractionError(
            f"{block.sheet_name!r} block {block.index}: this block extracts neither "
            f"{' nor '.join(repr(f) for f in missing)} nor {bounds.label!r}, so the band "
            f"bounds cannot be produced. Declare the two bound columns in the extraction "
            f"row, or a label they can be read off."
        )

    records, derived = list(block.records), ()
    if missing:
        records = []
        for record in block.records:
            lower, upper = parse_band(record.values.get(bounds.label), record.source_ref)
            values = dict(record.values)
            values.setdefault(bounds.lower, lower)
            values.setdefault(bounds.upper, upper)
            records.append(Record(record.source_ref, values, record.confidence))
        derived = tuple(missing)

    pairs = [(r.values.get(bounds.label) or _span(r, bounds),
              r.values.get(bounds.lower), r.values.get(bounds.upper))
             for r in records]
    return records, derived, continuity(pairs)


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
        return "both", (tuple(occ), tuple(cov))
    if occ:
        return "occupancy", tuple(occ)
    if cov:
        return "cover", tuple(cov)
    if seg:
        return "segment", tuple(seg)
    if total in have:
        return "total", (total,)
    return "", ()


def _refuse_unused_levels(block: Block, targets, total: str, kind: str, labels) -> None:
    """A reported column the bridge does not consume must never be dropped in silence.

    Two columns describing the same book on different axes is not a level the tool can
    read — and quietly using one and discarding the other would lose money without
    leaving a trace, which is the one thing this step must not do.
    """
    used = set(labels[0] + labels[1]) if kind == "both" else set(labels)
    unused = [f for f in block.numeric_fields
              if f not in used and f != total and f not in set(targets)]
    if unused:
        raise ExtractionError(
            f"{block.sheet_name!r} block {block.index}: the block reports "
            f"{', '.join(unused)}, which the split does not consume at the level it "
            f"chose ({kind or 'none'}). Either that column belongs to a level sheet 00 "
            "does not declare, or two levels arrived at once — both need saying out "
            "loud rather than one of them being dropped"
        )


def _split(block: Block, spec: Step2Spec, records, nomenclature, blocks=None):
    """Bridge what arrived onto the dataset's target cells — spec §2.6, S20.

    Two steps, and keeping them apart is what makes every case one case: **expand** what
    the source reports onto the section's occupancy × cover grid, then **project** that
    grid onto the axes the dataset actually declares. Earthquake declares both, so
    nothing is projected away; windstorm declares occupancy alone, so the cover axis is
    summed out — which still uses the cover information rather than discarding it.

    A split redistributes and never creates: the record total is unchanged.
    """
    rule = spec.split
    if rule is None or nomenclature is None:
        return records, (), [], False, ()

    axes = nomenclature.axes_for(block.dataset.key)
    targets = nomenclature.buckets_for(block.dataset.key)
    if not axes or not targets:
        return records, (), [], False, ()
    if all(t in block.fields for t in targets):
        # The finished grid arrived. Any *other* level column describes the same book a
        # second way, and choosing between them silently is not the tool's to do.
        _refuse_unused_levels(block, targets, rule.total, "reported", targets)
        return records, (), [], False, ()

    occupancy = tuple(axes[0].categories)
    target_covers = tuple(axes[1].categories) if len(axes) > 1 else ()
    segments = tuple(nomenclature.segment_categories())

    bridge = build_bridge(nomenclature, blocks or [], block.section,
                          occupancy, target_covers or nomenclature.cover_categories())
    if bridge is None:
        return _declared_split(block, spec, records, nomenclature, targets, axes)

    kind, labels = _reported_level(block, occupancy, bridge.covers, segments, rule.total)
    _refuse_unused_levels(block, targets, rule.total, kind, labels)
    if not kind:
        raise ExtractionError(
            f"{block.sheet_name!r} block {block.index}: the source reports none of the "
            f"{len(targets)} target cell(s), neither margin, no segmentation and no "
            f"{rule.total} — there is no level to build from, and the tool will not "
            "invent one"
        )

    out, notes = [], []
    for record in records:
        values = dict(record.values)
        grid = _expand(record, kind, labels, bridge, occupancy)
        for name, cell in _project(targets, axes, occupancy, bridge.covers).items():
            values[name] = sum(grid.get(k, 0.0) for k in cell)
        out.append(Record(record.source_ref, values, record.confidence))

    notes.extend(_split_notes(kind, labels, bridge, targets, axes))
    created = tuple(t for t in targets if t not in block.fields)
    return out, created, notes, kind == "total", created


def _expand(record, kind: str, labels, bridge, occupancy) -> dict:
    """One record's amounts, spread over the section's occupancy × cover grid."""
    if kind == "both":
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
        source = "total" if kind == "total" else kind
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
        "both": "both margins reported per zone, fitted to the grid from {src} so that "
                "neither margin moves",
        "occupancy": "occupancy reported per zone; the cover mix of each comes from {src}",
        "cover": "cover reported per zone; the occupancy mix of each comes from {src} — "
                 "which is why residential BI stays near nil rather than taking a flat share",
        "segment": "reported as {shown}, translated onto the target cells through {src} "
                   "and the declared occupancy convention",
        "total": "one figure per zone, spread over the target cells by the joint "
                 "distribution of {src}",
    }[kind]
    shown = ", ".join(labels if kind != "both" else labels[0] + labels[1])
    notes = [
        f"{len(targets)} target cell(s) built: "
        + told.format(src=bridge.source, shown=shown)
        + " (value-adding — an assumption about the zone, not about the book)"
    ]
    if kind == "segment" and bridge.conventions:
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


def _declared_split(block, spec, records, nomenclature, targets, axes):
    """No 08 for this section — fall back to ratios typed into ⟦SPLITS⟧ — spec §2.6."""
    rule = spec.split
    if rule.total not in block.fields:
        raise ExtractionError(
            f"{block.sheet_name!r} block {block.index}: section {block.section!r} carries "
            f"no 08 split table and the block reports no {rule.total}, so there is "
            "nothing to build the target cells from"
        )

    shares, sources = {}, []
    for name in targets:
        factor, origin = 1.0, []
        for axis, category in zip(axes, _cells_of(name, axes)):
            declared = {r.category: r for r in
                        nomenclature.splits_for(block.dataset.key, axis.name)}
            if category not in declared:
                raise ExtractionError(
                    f"{block.sheet_name!r} block {block.index}: no 08 table and sheet 00 "
                    f"⟦SPLITS⟧ declares no share for {category!r} on axis "
                    f"{axis.name!r} — the tool will not invent one"
                )
            factor *= declared[category].share
            if declared[category].source:
                origin.append(declared[category].source)
        shares[name] = factor
        sources.extend(origin)

    out = []
    for record in records:
        values = dict(record.values)
        base = _amount(record, rule.total)
        for name, share in shares.items():
            values[name] = base * share
        out.append(Record(record.source_ref, values, record.confidence))

    origin = "; ".join(dict.fromkeys(sources))
    created = tuple(t for t in targets if t not in block.fields)
    return out, created, [
        f"no 08 split table for section {block.section!r}; {len(targets)} target cell(s) "
        f"built from ratios declared in sheet 00 ⟦SPLITS⟧"
        + (f" [{origin}]" if origin else "")
        + " (value-adding — an assumption, not a reading)",
        "the split redistributes and leaves every zone total unchanged",
    ], True, created


def _cells_of(name: str, axes) -> list[str]:
    """Split a target cell name back into one category per axis."""
    if len(axes) == 1:
        return [name]
    for first in axes[0].categories:
        if name.startswith(f"{first} ") and name[len(first) + 1:] in axes[1].categories:
            return [first, name[len(first) + 1:]]
    return [name]


def _complete(block: Block, spec: Step2Spec, records, nomenclature, extra=()):
    """Add the catalogue members the source is silent about, with 0 — spec §9.2.1 S18.

    A cat aggregate lists the zones the cedent has exposure in. The zones with none are
    the ones worth seeing: absent reads as "no data", zero reads as "nothing there".
    Adding zeros cannot move a total, so this stays value-preserving.
    """
    rule = spec.complete
    if rule is None or nomenclature is None or rule.key not in block.fields:
        return records, []

    declared = block.attributes.get(rule.catalogue_attribute)
    if declared is None:
        return records, []
    catalogue = nomenclature.zones_for(str(declared.value))
    if not catalogue:
        raise ExtractionError(
            f"{block.sheet_name!r} block {block.index}: "
            f"{rule.catalogue_attribute} = {declared.value!r}, which sheet 00 ⟦ZONES⟧ "
            "does not list — the members with no exposure cannot be shown without it"
        )

    present = {norm(r.values.get(rule.key)) for r in records}
    filled = []
    for member in catalogue:
        if norm(member) in present:
            continue
        values = {rule.key: member}
        values.update({m: 0.0 for m in (*block.measure_fields, *extra)})
        records.append(Record("—", values))
        filled.append(member)
    return records, filled


def _identity(block: Block, spec: Step2Spec, records, nomenclature=None,
              from_total: bool = False):
    """Reconcile a declared total against its parts — spec §9.2.1 S19.

    Derived where the source omits it, checked where the source supplies it, and left
    alone where the parts are missing — the same three shapes as the band bounds.

    ``from_total`` says the parts were themselves computed *from* the total by the split.
    Then the sum agrees by construction and there is nothing to check, so the block says
    that rather than reporting a passed test it could never have failed.
    """
    rule = spec.identity
    if rule is None:
        return records, (), None

    declared = rule.parts or (
        nomenclature.buckets_for(block.dataset.key) if nomenclature else ()
    )
    if not declared:
        return records, (), None

    known = set(block.fields) | set(records[0].values) if records else set(block.fields)
    parts = [p for p in declared if p in known]
    if not parts:
        return records, (), (
            f"{rule.total} stands alone: the source declares none of "
            f"{declared[0]} … {declared[-1]}, so the split cannot be reconciled"
        )

    if from_total:
        return records, (), (
            f"{rule.total} was the source of the {len(parts)} part(s), so they sum back "
            "to it by construction — nothing here is a check"
        )

    if rule.total in block.fields:
        off = []
        for record in records:
            total = record.values.get(rule.total)
            summed = sum(record.values[p] for p in parts
                         if isinstance(record.values.get(p), (int, float)))
            if isinstance(total, (int, float)) and abs(total - summed) > 0.5:
                off.append(f"{record.values.get(rule.total, '')}"
                           f"{record.source_ref}: {total:,.0f} vs {summed:,.0f}")
        note = (f"{rule.total} checked against {len(parts)} part(s): "
                + ("all agree" if not off else "DISAGREES — " + "; ".join(off[:5])))
        return records, (), note

    out = []
    for record in records:
        values = dict(record.values)
        values[rule.total] = sum(values[p] for p in parts
                                 if isinstance(values.get(p), (int, float)))
        out.append(Record(record.source_ref, values, record.confidence))
    return out, (rule.total,), (
        f"{rule.total} derived as the sum of {len(parts)} declared part(s) — the source "
        "does not supply it (value-adding, but arithmetic only)"
    )


def _cumulative(records, spec: Step2Spec) -> dict[str, list[float | None]]:
    """Running share of each column's total — spec §9.2.1 S17."""
    out = {}
    for cum in spec.cumulative:
        total = sum(r.values[cum.field] for r in records
                    if isinstance(r.values.get(cum.field), (int, float)))
        running, column = 0.0, []
        for record in records:
            value = record.values.get(cum.field)
            if isinstance(value, (int, float)):
                running += float(value)
            column.append(running / total if total else None)
        out[cum.name] = column
    return out


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


def _resolve_group_by(block: Block, spec: AggregateSpec) -> str:
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


def _aggregate(block: Block, spec: AggregateSpec, blocks) -> AggregateTable:
    """Group and sum, spanning the declared year window — spec §9.2.1 S14."""
    from .crosschecks import split_label, years_in

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
    return AggregateTable(spec.title, group_by, spec.measures, rows,
                          sorted(zero_filled), spec.note)


def apply_step2(block: Block, spec: Step2Spec, nomenclature=None, blocks=None) -> Step2Result:
    missing = [f for f in spec.sort_by if f not in block.dataset.headers]
    if missing:
        raise ExtractionError(f"step 2 sorts by {missing}, which the dataset does not declare")

    actual_year = getattr(nomenclature, "actual_year", None)
    order = getattr(nomenclature, "period_order", None) or {}

    # Bounds first: the sort depends on them, and they may have to be read off the label.
    source, bound_fields, gaps = _derive_bounds(block, spec)
    # The split comes first: the identity has nothing to reconcile until the buckets
    # exist, and a zone added as 0 splits to zeros whichever way round it is done.
    source, split_fields, split_notes, from_total, assumed = _split(
        block, spec, source, nomenclature, blocks)
    source, total_field, identity_note = _identity(block, spec, source,
                                                   nomenclature, from_total)
    derived_fields = bound_fields + split_fields + total_field
    source, filled = _complete(block, spec, source, nomenclature, split_fields)

    records = sorted(source,
                     key=lambda r: _sort_key(r, spec.sort_by, order, spec.numeric_sort),
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

    computed.update(_cumulative(records, spec))

    sort_note = (f"sorted by {', '.join(spec.sort_by)} "
                 f"{'ascending' if spec.ascending else 'descending'}, "
                 + ("numeric" if spec.numeric_sort else "natural alphanumeric"))
    if order and not spec.numeric_sort:
        ranked = " → ".join(s for s, _ in sorted(order.items(), key=lambda kv: kv[1]))
        sort_note += f", period order {ranked}"

    notes = []
    if bound_fields:
        # Only the band bounds are read off a label; a derived Total speaks for itself
        # through identity_note below.
        shown = "; ".join(
            f"{r.values.get(spec.bounds.label)!r} → "
            f"{'' if r.values.get(spec.bounds.lower) is None else format(r.values[spec.bounds.lower], ',.0f')}"
            f" … "
            f"{'open' if r.values.get(spec.bounds.upper) is None else format(r.values[spec.bounds.upper], ',.0f')}"
            for r in records[:3]
        )
        notes.append(
            f"{' and '.join(bound_fields)} read off {spec.bounds.label} — the source "
            f"declares no such column (value-adding): {shown}"
            + (" …" if len(records) > 3 else "")
        )

    available = set(block.fields) | set(derived_fields)
    shown_columns = [h for h in block.dataset.headers if h in available]
    notes.extend(split_notes)
    if identity_note:
        notes.append(identity_note)
    if filled:
        notes.append(
            f"{len(filled)} declared {spec.complete.key.lower()}(s) with no entry in the "
            f"source, shown as 0: {', '.join(filled)} (value-preserving)"
        )
    notes.append(f"{sort_note} (value-preserving)")
    notes.append(
        f"columns ordered as declared in sheet 00: "
        f"{' | '.join(shown_columns)} (value-preserving)"
    )
    for calc in spec.calculations:
        readable = FIELD.sub(lambda m: m.group(1), calc.expression).replace("/", " / ")
        notes.append(f"calculated {calc.name} = {readable} (value-adding)")
    for cum in spec.cumulative:
        notes.append(f"calculated {cum.name} = running {cum.field} ÷ total {cum.field} "
                     "(value-adding)")
    for finding in gaps or []:
        notes.append(f"BAND CONTINUITY: {finding}")

    aggregates = [_aggregate(block, a, blocks) for a in spec.aggregates]
    for table in aggregates:
        notes.append(
            f"aggregated: {table.title} — {table.group_by} × "
            f"{', '.join(table.measures)} (value-preserving in total)"
        )
        if table.zero_filled:
            notes.append(
                f"{table.title}: groups with no record shown as 0: "
                f"{', '.join(table.zero_filled)}"
            )

    result = Step2Result(
        block=block, spec=spec, records=records, computed=computed, notes=notes,
        figures=_derived_figures(block, spec, actual_year), aggregates=aggregates,
        derived_fields=derived_fields, assumed_fields=assumed,
    )
    if gaps:
        block.hypotheses.append(
            Hypothesis(
                id="", dataset_key=block.dataset.key, attribute="Band continuity",
                value=f"{len(gaps)} finding(s)", confidence=Confidence.OPEN,
                source="tool", note="; ".join(gaps),
            )
        )
        from .extract import _number_hypotheses
        _number_hypotheses(block)

    # Every grouping of the same records must reach the same total.
    for table in aggregates:
        for measure, total in table.totals().items():
            if abs(total - result.totals().get(measure, 0.0)) > 1e-9:
                raise ExtractionError(
                    f"{table.title!r}: the aggregate total for {measure!r} is {total}, "
                    f"but the detail totals {result.totals().get(measure)} — "
                    "grouping lost records"
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
