"""Band bounds. Specification_v1.md §2.4, §9.2.1 S16.

A band arrives in three shapes and step 2 must end with both bounds in every one of them.
Reading ``1`` and ``10,000`` out of ``"1-10,000"`` is interpretation, so it belongs here
and never to step 1 — which is why the derived records are copies.
"""

from __future__ import annotations

from dataclasses import dataclass

from ..constants import NOTE_EXAMPLES
from ..model import ExtractionError, Record
from . import Context


def span(record, lower: str, upper: str) -> str:
    """A band's name when the source gave no label — its bounds are its identity."""
    low, high = record.values.get(lower), record.values.get(upper)
    return (f"{'open' if low is None else format(low, ',.0f')} – "
            f"{'open' if high is None else format(high, ',.0f')}")


@dataclass(frozen=True)
class DeriveBounds:
    """Make sure step 2 always has both bounds, however the source supplied them.

    ==================================== ==================================================
    two numeric columns                  extracted; nothing is worked out
    a label only (``1 - 1,000,000``)     the bounds are read off it and reported
    two columns and no label             extracted; the band is named by its own bounds
    ==================================== ==================================================

    What must *not* happen is neither: a block extracting no bounds and no label cannot
    produce a profile, and says so rather than sorting on nothing.
    """

    label: str
    lower: str
    upper: str

    def apply(self, ctx: Context) -> None:
        from ..bands import continuity, parse_band

        block = ctx.block
        has_label = self.label in block.fields
        missing = [f for f in (self.lower, self.upper) if f not in block.fields]

        if missing and not has_label:
            raise ExtractionError(
                f"{block.sheet_name!r} block {block.index}: this block extracts neither "
                f"{' nor '.join(repr(f) for f in missing)} nor {self.label!r}, so the "
                f"band bounds cannot be produced. Declare the two bound columns in the "
                f"extraction row, or a label they can be read off."
            )

        if missing:
            records = []
            for record in ctx.records:
                low, high = parse_band(record.values.get(self.label), record.source_ref)
                values = dict(record.values)
                values.setdefault(self.lower, low)
                values.setdefault(self.upper, high)
                records.append(Record(record.source_ref, values, record.confidence))
            ctx.records = records
            ctx.derive(*missing)

            shown = "; ".join(
                f"{r.values.get(self.label)!r} → "
                f"{'' if r.values.get(self.lower) is None else format(r.values[self.lower], ',.0f')}"
                f" … "
                f"{'open' if r.values.get(self.upper) is None else format(r.values[self.upper], ',.0f')}"
                for r in ctx.records[:NOTE_EXAMPLES]
            )
            ctx.notes.append(
                f"{' and '.join(missing)} read off {self.label} — the source declares no "
                f"such column (value-adding): {shown}"
                + (" …" if len(ctx.records) > NOTE_EXAMPLES else "")
            )

        pairs = [(r.values.get(self.label) or span(r, self.lower, self.upper),
                  r.values.get(self.lower), r.values.get(self.upper))
                 for r in ctx.records]
        ctx.flags["band_gaps"] = continuity(pairs) or []

    def not_summable(self) -> set[str]:
        """A band bound is a number but not a quantity — spec §2.4."""
        return {self.lower, self.upper}


@dataclass(frozen=True)
class BandContinuity:
    """Report the gaps :class:`DeriveBounds` found — spec §2.4.

    Separate from the derivation because it is a *finding*, not a step: both conventions
    are in use (shared boundary, gapless integers) and only a band genuinely missing from
    the middle of a profile is worth saying out loud.
    """

    def apply(self, ctx: Context) -> None:
        gaps = ctx.flags.get("band_gaps", ())
        for finding in gaps:
            ctx.notes.append(f"BAND CONTINUITY: {finding}")
        if gaps:
            from ..extract import _number_hypotheses
            from ..model import Confidence, Hypothesis

            ctx.block.hypotheses.append(Hypothesis(
                id="", dataset_key=ctx.block.dataset.key, attribute="Band continuity",
                value=f"{len(gaps)} finding(s)", confidence=Confidence.OPEN,
                source="tool", note="; ".join(gaps),
            ))
            _number_hypotheses(ctx.block)

    def not_summable(self) -> set[str]:
        return set()
