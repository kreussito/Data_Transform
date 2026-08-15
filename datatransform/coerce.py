"""Type coercion. Specification_v1.md §8.4.

Applied at extraction, because step 1's control sums must be arithmetically
meaningful — a sum over text silently under-counts.

Two rules govern every conversion here:

* **Empty stays empty.** ``None`` is an absence, not a zero. Zero is a claim about the
  data; conflating the two understates a loss ratio without leaving a trace.
* **Ambiguity is fatal.** ``"1.234"`` is 1234 under a German convention and 1.234 under
  an English one. Guessing is a 1000× error in a premium figure, so an ambiguous form
  stops the run instead.
"""

from __future__ import annotations

import re
from datetime import date, datetime

from .model import ExtractionError, FieldType

THOUSANDS_GROUP = re.compile(r"^\d{1,3}(?:([.,  ])\d{3})+$")
CLEAN = str.maketrans({" ": "", " ": "", " ": "", "'": "", "’": ""})


class Coercion:
    """A conversion applied on reading, so step 1 can report it.

    ``routine`` marks the declared type simply being applied to a well-formed cell —
    ``2021`` read as ``"2021"``. Those are already covered by the block's "Types
    applied" line, so only *notable* conversions (a number arriving as text, a date
    where a year was expected) are surfaced to the reviewer. Both kinds reach the
    debug log.
    """

    __slots__ = ("field", "source_ref", "before", "after", "note", "routine")

    def __init__(self, field, source_ref, before, after, note, routine=False):
        self.field = field
        self.source_ref = source_ref
        self.before = before
        self.after = after
        self.note = note
        self.routine = routine

    def __repr__(self):
        return f"<Coercion {self.field}@{self.source_ref} {self.before!r}→{self.after!r}>"


def to_text(value, field: str, ref: str, log: list) -> str | None:
    if value is None or (isinstance(value, str) and not value.strip()):
        return None

    if isinstance(value, bool):
        raise ExtractionError(f"{field} at {ref}: boolean {value!r} cannot be read as text")

    if isinstance(value, (datetime, date)):
        out = str(value.year)
        log.append(Coercion(field, ref, value, out,
                            "cell holds a date; its year was taken"))
        return out

    if isinstance(value, float):
        if not value.is_integer():
            raise ExtractionError(
                f"{field} at {ref}: {value!r} is not a whole number and cannot be read "
                f"as a {field} label"
            )
        out = str(int(value))
        log.append(Coercion(field, ref, value, out, "numeric cell read as text", routine=True))
        return out

    if isinstance(value, int):
        out = str(value)
        log.append(Coercion(field, ref, value, out, "numeric cell read as text", routine=True))
        return out

    return str(value).strip()


def to_number(value, field: str, ref: str, log: list) -> float | None:
    if value is None or (isinstance(value, str) and not value.strip()):
        return None

    if isinstance(value, bool):
        raise ExtractionError(f"{field} at {ref}: boolean {value!r} cannot be read as a number")

    if isinstance(value, (int, float)):
        return float(value)

    if isinstance(value, (datetime, date)):
        raise ExtractionError(f"{field} at {ref}: a date cannot be read as a number")

    raw = str(value).strip()
    text = raw.translate(CLEAN)

    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1]
    text = text.lstrip("+")
    if text.startswith("-"):
        negative, text = True, text[1:]
    text = re.sub(r"[A-Za-z%€$£]+$", "", text).strip()

    if not text:
        return None

    number = _parse_decimal(text, field, ref, raw)
    out = -number if negative else number
    if raw != str(out):
        log.append(Coercion(field, ref, raw, out, "text cell read as a number"))
    return out


def _parse_decimal(text: str, field: str, ref: str, raw: str) -> float:
    has_dot, has_comma = "." in text, "," in text

    if has_dot and has_comma:
        # The rightmost separator is the decimal point; the other groups thousands.
        decimal_sep = "." if text.rfind(".") > text.rfind(",") else ","
        thousands = "," if decimal_sep == "." else "."
        cleaned = text.replace(thousands, "").replace(decimal_sep, ".")
        return _to_float(cleaned, field, ref, raw)

    if not has_dot and not has_comma:
        return _to_float(text, field, ref, raw)

    sep = "." if has_dot else ","
    parts = text.split(sep)

    if len(parts) > 2:                                  # 1.234.567 — grouping
        return _to_float(text.replace(sep, ""), field, ref, raw)

    head, tail = parts
    if len(tail) != 3:                                  # 12.5 or 1.2345 — a decimal point
        return _to_float(text.replace(sep, "."), field, ref, raw)

    # Exactly three trailing digits, one separator: 1.234 is genuinely ambiguous.
    if THOUSANDS_GROUP.match(text) and len(head) <= 3:
        raise ExtractionError(
            f"{field} at {ref}: {raw!r} is ambiguous — {sep!r} could be a thousands "
            f"separator ({text.replace(sep, '')}) or a decimal point "
            f"({text.replace(sep, '.')}). Format the cell as a number in Excel, or "
            f"state the convention, rather than have it guessed."
        )
    return _to_float(text.replace(sep, ""), field, ref, raw)


def _to_float(text: str, field: str, ref: str, raw: str) -> float:
    try:
        return float(text)
    except ValueError:
        raise ExtractionError(
            f"{field} at {ref}: {raw!r} cannot be read as a number"
        ) from None


COERCERS = {
    FieldType.TEXT: to_text,
    FieldType.NUMBER: to_number,
}


def coerce(value, field_type: FieldType, field: str, ref: str, log: list):
    return COERCERS[field_type](value, field, ref, log)
