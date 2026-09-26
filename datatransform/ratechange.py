"""What the cedent claims against what the figures imply. Specification_v1.md §10.5.

§10.4 works out an **implied** rate change from the two things a submission cannot fake:
premium and exposure. ``09`` carries the **claimed** one — the figure a cedent, a broker
or the underwriter's own note puts on the renewal. Neither is the truth. Where they agree
the renewal rests on something; where they do not, one of them is wrong and the gap is
the conversation.

**Nothing here fails a run.** The two numbers measure the book differently and drift for
honest reasons, so a gap is a *finding with a question*, exactly as §2.6 treats two views
of the same split. It is written under the rate sheet with somewhere to answer.

**A scope may be several sections.** A nat cat programme is usually quoted as one rate
for earthquake and windstorm together, so ``09`` blocks declare ``Scope = Earthquake +
Hurricane`` and the implied side has to be combined to match. That combination is the one
piece of arithmetic here that is not obvious, and §10.5 spells out why:

* **Premium adds.** Earthquake premium plus hurricane premium is the nat cat premium.
  Nothing is counted twice.
* **Exposure does not.** A coastal hotel sits in the earthquake aggregate *and* in the
  windstorm one. Their sum is not a portfolio figure and must never be read as one.

But the implied rate uses only *growth*, never a level — and the growth of the sum is the
exposure-weighted mean of the two growths, in which a stable double count cancels. So the
combination is legitimate precisely because it never looks at the total.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .constants import (
    A_SCOPE,
    A_SOURCE,
    F_RATE_CHANGE,
    F_YEAR,
    NO_BASIS,
    RATE_CLAIM_DEFAULT,
    RATE_CLAIM_WARNING,
    ROLE_RATE,
    SCOPE_JOIN,
    WARNING,
    WITHIN_LIMITS,
    read_share,
)
from .model import ExtractionError
from .nomenclature import norm


def threshold_of(nomenclature) -> float:
    """⟦GLOBAL⟧ ``Rate claim warning`` — in **percentage points**, not as a ratio.

    The two sides are themselves rate changes, so the meaningful distance between them is
    a difference: +4% claimed against −3.4% implied is a gap of 7.4 points, and calling
    that "218%" would be arithmetic without meaning.
    """
    return read_share(nomenclature, RATE_CLAIM_WARNING, RATE_CLAIM_DEFAULT)


def scope_of(block, nomenclature) -> tuple[str, ...]:
    """The sections one ``09`` block speaks for — spec §10.5.

    ``Scope = Earthquake + Hurricane`` names two; an absent ``Scope`` falls back to the
    block's own ``Section``, which is the ordinary case. A name that ⟦SECTIONS⟧ does not
    declare is fatal: a rate compared against the wrong book is worse than no comparison.
    """
    attr = block.attributes.get(A_SCOPE)
    raw = norm(attr.value) if attr else (block.section or "")
    if not raw:
        return ()

    declared = {norm(s.name).casefold(): s.name
                for s in (getattr(nomenclature, "sections", None) or [])}
    out = []
    for part in (p.strip() for p in raw.split(SCOPE_JOIN)):
        if not part:
            continue
        match = declared.get(norm(part).casefold())
        if match is None:
            raise ExtractionError(
                f"{block.sheet_name!r} block {block.index}: {A_SCOPE} names {part!r}, "
                f"which ⟦SECTIONS⟧ does not declare "
                f"({', '.join(declared.values()) or 'no sections declared'})"
            )
        out.append(match)
    return tuple(dict.fromkeys(out))


@dataclass
class RateClaim:
    """One claimed rate change, and the implied one it is held against — spec §10.5."""

    scope: tuple[str, ...]
    year: str
    claimed: float | None
    sheets: tuple[str, ...]
    threshold: float
    source: str = ""
    premium_growth: float | None = None
    exposure_growth: float | None = None
    parts: list[tuple[str, float | None, float | None]] = field(default_factory=list)
    skipped: str = ""

    @property
    def title(self) -> str:
        return f"Claimed rate change against implied — {' + '.join(self.scope)}"

    @property
    def combined(self) -> bool:
        return len(self.scope) > 1

    @property
    def implied(self) -> float | None:
        if self.premium_growth is None or self.exposure_growth is None:
            return None
        if self.exposure_growth == -1:
            return None
        return (1 + self.premium_growth) / (1 + self.exposure_growth) - 1

    @property
    def gap(self) -> float | None:
        """In percentage points — both sides are already changes."""
        if self.claimed is None or self.implied is None:
            return None
        return self.claimed - self.implied

    @property
    def status(self) -> str:
        gap = self.gap
        if gap is None:
            return NO_BASIS
        return WARNING if abs(gap) > self.threshold else WITHIN_LIMITS

    @property
    def question(self) -> str:
        gap = self.gap
        if gap is None:
            return (
                "QUESTION FOR THE UNDERWRITER: there is no implied rate to hold this "
                "against — the exposure or the premium for this scope is missing, so the "
                "claim stands unchecked. Is that expected?"
            )
        direction = "above" if gap > 0 else "below"
        return (
            f"QUESTION FOR THE UNDERWRITER: the claimed rate change is "
            f"{abs(gap):.1%} {direction} what premium and exposure imply. Either the "
            "claim is measured on a different basis — risk-adjusted, renewal-only, net "
            "of commission — or one of the two figures is wrong. Which is it?"
        )


def _renewal_record(block, year: int | None):
    """The record for N+1, which is the year a renewal turns on."""
    if year is None:
        return None
    wanted = str(year + 1)
    for record in block.records:
        if norm(record.values.get(F_YEAR)) == wanted:
            return record
    return None


def _combine(tables) -> tuple[float | None, float | None, list]:
    """Premium adds; exposure growth is weighted. See the module docstring."""
    parts = [(t.section, t.premium_growth, t.exposure_growth) for t in tables]

    premium_from = sum(t.premium_from[1] for t in tables if t.premium_from)
    premium_to = sum(t.premium_to[1] for t in tables if t.premium_to)
    premium = (premium_to / premium_from - 1) if premium_from else None

    weighted, weight = 0.0, 0.0
    for table in tables:
        growth = table.exposure_growth
        base = table.versions[0].exposure if table.versions else 0.0
        if growth is not None and base:
            weighted += growth * base
            weight += base
    exposure = weighted / weight if weight else None
    return premium, exposure, parts


def rate_claims(nomenclature, blocks, growth) -> list[RateClaim]:
    """One entry per ``09`` block — spec §10.5."""
    if isinstance(blocks, dict):
        blocks = list(blocks.values())

    n = nomenclature.actual_year
    threshold = threshold_of(nomenclature)
    out = []

    for block in [b for b in blocks if b.dataset.role == ROLE_RATE]:
        scope = scope_of(block, nomenclature)
        record = _renewal_record(block, n)
        claimed = record.values.get(F_RATE_CHANGE) if record else None
        source = block.attributes.get(A_SOURCE)

        claim = RateClaim(
            scope=scope, year=str((n or 0) + 1),
            claimed=claimed if isinstance(claimed, (int, float)) else None,
            sheets=(block.sheet_name,), threshold=threshold,
            source=norm(source.value) if source else "",
        )

        if not scope:
            claim.skipped = (
                f"the block declares neither {A_SCOPE} nor a Section, so there is no "
                "book to hold the claim against"
            )
            out.append(claim)
            continue
        if claim.claimed is None:
            claim.skipped = (
                f"no {F_RATE_CHANGE} for {claim.year} — the renewal year is the one a "
                "rate claim is about, and this block does not carry it"
            )
            out.append(claim)
            continue

        tables = [t for t in growth if t.section in scope and not t.skipped]
        if not tables:
            claim.skipped = (
                f"no exposure growth for {' + '.join(scope)}: §10.4 produced no table, so "
                "the claim cannot be checked — it is shown unchecked rather than dropped"
            )
            out.append(claim)
            continue

        premium, exposure, parts = _combine(tables)
        claim.premium_growth, claim.exposure_growth, claim.parts = premium, exposure, parts
        out.append(claim)
    return out


def claims_for_sheet(claims, sheet_name: str) -> list[RateClaim]:
    return [c for c in claims if sheet_name in c.sheets]
