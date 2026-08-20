"""Exposure growth against premium growth. Specification_v1.md §10.4.

A cat aggregate arrives in three versions — the nine-month estimate of the expiring year
and projections for the first and last day of the renewal year. No single version is
interesting; the movement between them is, and only against the premium movement.

The figure this block exists to produce is the **implied rate change**::

    (1 + premium growth) ÷ (1 + exposure growth) − 1

Premium up 9% carried on 13% more exposure is a *rate cut* of about 3.5%, however the
premium line reads on its own. That is the number a renewal turns on, and it is also the
bridge to ``09. Rate Development``: if the cedent claims +4% and this says −3.5%, one of
the two is wrong.

**Growth is never an error.** Portfolios shrink for good reasons — a cedent drops a
segment, a currency moves, a large scheme leaves. So nothing here fails a run. An
implied rate change beyond the declared threshold is a *warning*, raised so the renewal
is priced knowingly, exactly as §10.3 treats an unusual loss share.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .constants import (
    EXPOSURE_ROLES,
    NO_BASIS,
    RATE_CHANGE_DEFAULT,
    RATE_CHANGE_WARNING,
    UNRANKED_PERIOD,
    WARNING,
    WITHIN_LIMITS,
    A_PERIOD,
    F_EPI,
    F_PREMIUM,
    F_TOTAL,
    ROLE_EPI,
    ROLE_HISTORY,
    read_share,
)
from .crosschecks import find_block, match_record
from .nomenclature import norm

EXPOSURE_FIELD = F_TOTAL
PERIOD_ATTRIBUTE = A_PERIOD
THRESHOLD_ATTRIBUTE = RATE_CHANGE_WARNING
DEFAULT_THRESHOLD = RATE_CHANGE_DEFAULT

OK = WITHIN_LIMITS


def threshold_of(nomenclature) -> float:
    """Declared in ⟦GLOBAL⟧ — the underwriter's number, not the tool's."""
    return read_share(nomenclature, THRESHOLD_ATTRIBUTE, DEFAULT_THRESHOLD)


@dataclass
class Version:
    """One version of the aggregate."""

    period: str
    as_at: str
    exposure: float
    change: float | None = None       # against the version before it


@dataclass
class GrowthTable:
    """One cat section's exposure against its premium — spec §10.4."""

    section: str
    role: str
    sheets: tuple[str, ...]
    threshold: float
    versions: list[Version] = field(default_factory=list)
    premium_from: tuple[str, float] | None = None      # (label, value) in year N
    premium_to: tuple[str, float] | None = None        # (label, value) in year N+1
    skipped: str = ""
    note: str = ""                                     # ordering the reader should check

    @property
    def title(self) -> str:
        return f"Exposure and premium growth — {self.section}"

    @property
    def exposure_growth(self) -> float | None:
        """First version to last — the span the implied rate is measured over."""
        if len(self.versions) < 2 or not self.versions[0].exposure:
            return None
        return self.versions[-1].exposure / self.versions[0].exposure - 1

    @property
    def premium_growth(self) -> float | None:
        if not self.premium_from or not self.premium_to or not self.premium_from[1]:
            return None
        return self.premium_to[1] / self.premium_from[1] - 1

    @property
    def implied_rate_change(self) -> float | None:
        exposure, premium = self.exposure_growth, self.premium_growth
        if exposure is None or premium is None or exposure == -1:
            return None
        return (1 + premium) / (1 + exposure) - 1

    @property
    def status(self) -> str:
        rate = self.implied_rate_change
        if rate is None:
            return NO_BASIS
        return WARNING if abs(rate) > self.threshold else OK


def _undeclared_suffixes(versions, order) -> list[str]:
    """Suffixes ⟦PERIOD ORDER⟧ does not rank — spec §9.2 S13.

    An unranked suffix sorts last, and two of them sort alphabetically against each
    other. That is a defensible default for a table nobody reads in order, but §10.4
    measures growth from the *first* version to the *last*: get the sequence wrong and
    the span is wrong, silently and plausibly. So it is said out loud.
    """
    ranked = {norm(s).casefold() for s in order}
    out = []
    for block in versions:
        label = _period_of(block)
        suffix = label.split(" ", 1)[1] if " " in label else ""
        if suffix and norm(suffix).casefold() not in ranked:
            out.append(suffix)
    return sorted(dict.fromkeys(out))


def _period_of(block) -> str:
    attr = block.attributes.get(PERIOD_ATTRIBUTE)
    return norm(attr.value) if attr else ""


def _exposure(carrier) -> float:
    """Total the exposure of anything holding records — a block or a step-2 result."""
    return sum(r.values[EXPOSURE_FIELD] for r in carrier.records
               if isinstance(r.values.get(EXPOSURE_FIELD), (int, float)))


def growth_tables(nomenclature, blocks, results=None) -> list[GrowthTable]:
    """One table per cat section that carries a versioned aggregate."""
    if isinstance(blocks, dict):
        blocks = list(blocks.values())

    from .transform import _period_key

    order = getattr(nomenclature, "period_order", None) or {}
    n = nomenclature.actual_year
    threshold = threshold_of(nomenclature)
    by_block = {id(r.block): r for r in (results or [])}

    out = []
    for section in getattr(nomenclature, "sections", None) or []:
        for role in EXPOSURE_ROLES:
            versions = [b for b in blocks
                        if b.dataset.role == role and b.section == section.name]
            if not versions:
                continue

            table = GrowthTable(
                section=section.name, role=role,
                sheets=tuple(dict.fromkeys(b.sheet_name for b in versions)),
                threshold=threshold,
            )

            undeclared = _undeclared_suffixes(versions, order)
            if undeclared:
                table.note = (
                    f"the suffix(es) {', '.join(undeclared)} have no rank in "
                    "⟦PERIOD ORDER⟧, so the versions were ordered alphabetically — "
                    "declare them if that is not the intended sequence"
                )

            unlabelled = [b for b in versions if not _period_of(b)]
            if unlabelled:
                table.skipped = (
                    f"{len(unlabelled)} block(s) declare no {PERIOD_ATTRIBUTE!r}, so the "
                    "versions cannot be put in order"
                )
                out.append(table)
                continue

            # Step 2 completes the zone list; its records are the ones to total, so the
            # figures here match what the reader sees in the block above.
            ordered = sorted(versions,
                             key=lambda b: (_period_key(_period_of(b), order)
                                            or ((UNRANKED_PERIOD, 0, ""),)))
            previous = None
            for block in ordered:
                # Prefer the step-2 records when they are to hand: completing the zone list
                # only adds zeros, but a *derived* Total exists nowhere else.
                result = by_block.get(id(block))
                exposure = _exposure(block if result is None else result)
                as_at = block.attributes.get("As at")
                change = (exposure / previous - 1) if previous else None
                table.versions.append(
                    Version(_period_of(block), norm(as_at.value) if as_at else "",
                            exposure, change)
                )
                previous = exposure

            _attach_premium(table, blocks, section.name, n)
            out.append(table)
    return out


def _attach_premium(table: GrowthTable, blocks, section: str, n: int | None) -> None:
    """Year N from 01, year N+1 from 02 — 01 has no forward figure, so 02 is the only one."""
    if n is None:
        return

    history = find_block(blocks, ROLE_HISTORY, section)
    if history is not None:
        record = match_record(history, str(n), n)
        if record is not None and isinstance(record.values.get(F_PREMIUM), (int, float)):
            table.premium_from = (str(n), float(record.values[F_PREMIUM]))

    projection = find_block(blocks, ROLE_EPI, section)
    if projection is not None:
        record = match_record(projection, str(n + 1), n)
        if record is not None and isinstance(record.values.get(F_EPI), (int, float)):
            table.premium_to = (f"{n + 1} EPI", float(record.values[F_EPI]))


def tables_for_sheet(tables, sheet_name: str) -> list[GrowthTable]:
    return [t for t in tables if sheet_name in t.sheets]
