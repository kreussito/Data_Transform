"""Reading a band label. Specification_v1.md §2.4.

A profile's bands arrive either as two numeric columns or as a single label:

===================== ============ ============
``1 - 10,000``        from 1       to 10,000
``10,001 – 20,000``   from 10,001  to 20,000
``> 1,000,000``       from 1,000,000
``up to 10,000``                   to 10,000
===================== ============ ============

Where the columns exist they are extracted and this module is not used. Where they do
not, step 2 reads the bounds off the label — never step 1, because reading ``1`` and
``10,000`` out of ``"1-10,000"`` is interpretation, and step 1 carries only what the
sheet says (spec §9.2 S11).

The discipline is the one already applied to decimals (§8.4): **unambiguous forms are
read, ambiguous ones are fatal.** A label nobody can read without guessing stops the run
and names itself; the escape hatch is always to declare the two numeric columns.

An open bound is an *absent* bound, never a zero. ``> 1,000,000`` has no upper bound, and
a profile's top band usually does not.
"""

from __future__ import annotations

import re

from .coerce import to_number
from .model import ExtractionError
from .nomenclature import norm

# Anything that separates two bounds: hyphen, en/em dash, "to", "..".
SEPARATOR = re.compile(r"\s*(?:--|-|–|—|\.\.|\bto\b|\bbis\b)\s*", re.I)

LOWER_ONLY = re.compile(r"^(?:>=?|≥|over|above|more than|ab|über)\s*(.+)$", re.I)
UPPER_ONLY = re.compile(r"^(?:<=?|≤|under|below|less than|up to|bis)\s*(.+)$", re.I)

# A bound that carries its own scale: "1 Mio", "10k", "2.5 m".
UNIT = re.compile(r"^(.*?)\s*(k|m|mio|mn|bn|tsd|thousand|million|billion)\.?$", re.I)
UNITS = {"k": 1e3, "tsd": 1e3, "thousand": 1e3,
         "m": 1e6, "mio": 1e6, "mn": 1e6, "million": 1e6,
         "bn": 1e9, "billion": 1e9}


class BandError(ExtractionError):
    """A label whose bounds cannot be read without guessing."""


def _bound(text: str, label: str, ref: str) -> float | None:
    """One side of a band. Empty, or a number — never a guess."""
    text = text.strip().strip("()[]").strip()
    if not text:
        return None

    factor = 1.0
    unit = UNIT.match(text)
    if unit and unit.group(1).strip():
        text, factor = unit.group(1).strip(), UNITS[unit.group(2).lower()]

    log: list = []
    try:
        value = to_number(text, "Band", ref, log)
    except ExtractionError as exc:
        raise BandError(
            f"band {label!r} at {ref}: {exc}. Declare 'Band from' and 'Band to' as "
            "columns rather than leaving it to be read off the label."
        ) from None
    if value is None:
        raise BandError(f"band {label!r} at {ref}: {text!r} is not a number")
    return value * factor


def parse_band(label, ref: str = "") -> tuple[float | None, float | None]:
    """``(lower, upper)`` — either may be ``None`` for an open bound."""
    text = norm(label)
    if not text:
        raise BandError(f"band label at {ref} is empty, so its bounds cannot be read")

    one_sided = LOWER_ONLY.match(text)
    if one_sided:
        return _bound(one_sided.group(1), text, ref), None
    one_sided = UPPER_ONLY.match(text)
    if one_sided:
        return None, _bound(one_sided.group(1), text, ref)

    # A leading minus belongs to the first bound, not to the separator.
    parts = SEPARATOR.split(text.lstrip("-"), maxsplit=1)
    if len(parts) == 2 and parts[0].strip() and parts[1].strip():
        lower = _bound(parts[0], text, ref)
        upper = _bound(parts[1], text, ref)
        if lower is not None and upper is not None and upper < lower:
            raise BandError(
                f"band {text!r} at {ref}: upper bound {upper:,.0f} is below the lower "
                f"bound {lower:,.0f}"
            )
        return lower, upper

    # A bare number is a band of one point — legitimate, and unambiguous.
    value = _bound(text, text, ref)
    return value, value


def bounds_for(records, label_field: str, lower_field: str, upper_field: str):
    """Bounds per record, read off the label. Returns a list of ``(lower, upper)``."""
    out = []
    for record in records:
        out.append(parse_band(record.values.get(label_field), record.source_ref))
    return out


def continuity(pairs) -> list[str]:
    """Gaps and overlaps between consecutive bands — spec §2.4.

    Both conventions are in use and both are correct::

        0 – 10,000 · 10,000 – 20,000      shared boundary
        1 – 10,000 · 10,001 – 20,000      gapless integers

    So a gap is reported only where the next band starts **more than one unit** above
    the previous band's end, and an overlap only where it starts below it. Neither
    convention is preferred, and neither is corrected.
    """
    findings = []
    for (a_label, _, a_to), (b_label, b_from, _) in zip(pairs, pairs[1:]):
        if a_to is None or b_from is None:
            continue                       # an open bound cannot be adjacent to anything
        if b_from < a_to:
            findings.append(f"{a_label!r} ends at {a_to:,.0f} but {b_label!r} starts at "
                            f"{b_from:,.0f} — the bands overlap")
        elif b_from > a_to + 1:
            findings.append(f"{a_label!r} ends at {a_to:,.0f} and {b_label!r} starts at "
                            f"{b_from:,.0f} — the range between them is in no band")
    return findings
