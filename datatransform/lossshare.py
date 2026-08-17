"""Declared losses against total incurred. Specification_v1.md §10.3.

`⟦RULES⟧` compares one figure with one other figure. This check cannot be written that
way, because both sides are *sums over several blocks*: a treaty may carry large losses
and cat losses at once, and a Nat Cat treaty carries one cat sheet per peril. So the
comparison is made over a **section group** — every section of one kind — and asks two
questions of each year:

* **Is it possible?** Individually reported losses cannot exceed the total incurred for
  the same year. R-01 and R-02 already check each loss dataset on its own; only the sum
  catches large *and* cat losses that are each plausible and jointly are not.
* **Is it ordinary?** A year where the declared losses make up more than the declared
  share of total incurred is driven by a handful of events rather than by attrition,
  which changes how it should be rated. That is a warning, not an error: it is a fact
  about the portfolio, not a mistake in the data.

The grouping is by section *kind*, which covers every case with one rule:

===================================== ====================================================
Engineering / Miscellaneous           one per-risk section; 03 and 04 summed if both exist
Fire only                             one per-risk section; 03 alone
Fire + Nat Cat                        per-risk group = Fire (03); cat group = EQ + Wind (04)
===================================== ====================================================
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .crosschecks import (
    _one_of,
    _same_section,
    match_record,
    match_records,
    scale_factor,
    years_in,
)
from .model import Block
from .nomenclature import norm

BASIS_ROLE = "01"
BASIS_FIELD = "Incurred Losses"
LOSS_ROLES = ("03", "04")
LOSS_FIELD = "Loss amount"

ROLE_LABEL = {"03": "Large losses (03)", "04": "Cat losses (04)"}

# The same tolerance R-01 and R-02 declare in sheet 00: figures rounded to thousands
# never satisfy exact equality, and a check that always fires gets ignored.
TOLERANCE = 1.0

DEFAULT_THRESHOLD = 0.20
THRESHOLD_ATTRIBUTE = "Loss share warning"

# Losses are compared with losses, so the premium basis is not among these.
PRECONDITIONS = ("Loss basis", "Share basis", "Currency")

# Deliberately not "OK": step 1 and step 2 already write an OK/MISMATCH control check,
# and one word must not mean two things on the same sheet.
OK, WARNING, EXCEEDS, NO_BASIS = "within limits", "WARNING", "EXCEEDS", "no basis"


@dataclass
class LossShareRow:
    """One year of one section group."""

    year: int
    by_role: dict[str, float]
    declared: float
    incurred: float | None
    reporting: int = 0                # sections that reported this year in 01
    expected: int = 0                 # sections in the group

    @property
    def partial(self) -> bool:
        """Not every section of the group reported a history row for this year."""
        return self.incurred is not None and 0 < self.reporting < self.expected

    @property
    def share(self) -> float | None:
        if self.incurred in (None, 0):
            return None
        return self.declared / self.incurred

    def status(self, threshold: float) -> str:
        if self.incurred is None:
            return NO_BASIS
        if self.declared > self.incurred + TOLERANCE:
            return EXCEEDS
        share = self.share
        if share is not None and share > threshold:
            return WARNING
        return OK


@dataclass
class LossShareTable:
    """One section group's check — spec §10.3."""

    kind: str
    sections: tuple[str, ...]
    roles: tuple[str, ...]
    sheets: tuple[str, ...]           # the 01 sheets this table is written on
    threshold: float
    rows: list[LossShareRow] = field(default_factory=list)
    skipped: str = ""                 # reason, where the group cannot be compared
    scale: str = ""

    @property
    def title(self) -> str:
        return f"Declared losses against total incurred — {self.kind} sections"

    def totals(self) -> tuple[float, float | None]:
        declared = sum(r.declared for r in self.rows)
        known = [r.incurred for r in self.rows if r.incurred is not None]
        return declared, (sum(known) if known else None)

    def statuses(self) -> list[str]:
        return [r.status(self.threshold) for r in self.rows]

    @property
    def worst(self) -> str:
        seen = self.statuses()
        for status in (EXCEEDS, NO_BASIS, WARNING):
            if status in seen:
                return status
        return OK


PERCENT = re.compile(r"^\s*([0-9.,]+)\s*%\s*$")


def threshold_of(nomenclature) -> float:
    """The warning threshold, declared in ⟦GLOBAL⟧ — spec §10.3.

    Accepts ``20%``, ``20`` and ``0.2`` alike: a share above 1 can only have been meant
    as a percentage. An undeclared threshold falls back to 20%, which is stated in the
    written table so the reader knows it was not their choice.
    """
    raw = (getattr(nomenclature, "globals", None) or {}).get(THRESHOLD_ATTRIBUTE)
    if raw is None or norm(raw) == "":
        return DEFAULT_THRESHOLD

    text = norm(raw)
    percent = PERCENT.match(text)
    try:
        value = float((percent.group(1) if percent else text).replace(",", ""))
    except ValueError:
        return DEFAULT_THRESHOLD
    if percent or value > 1:
        value /= 100.0
    return value


def _contributors(blocks, role: str, sections) -> tuple[dict[str, Block], list[str]]:
    """One block per section for this role — or a reason it could not be resolved.

    A section contributes once. `Intake_v1.xlsx` carries `01. History` and its
    transposed twin, which are the same figures in two shapes; summing both would
    double the portfolio.
    """
    resolved, ambiguous = {}, []
    for name in sections:
        candidates = [b for b in blocks
                      if b.dataset.role == role and _same_section(b, name)]
        if not candidates:
            continue
        block = _one_of(candidates)
        if block is None:
            ambiguous.append(
                f"{len(candidates)} blocks play role {role} in section {name!r}"
            )
            continue
        resolved[name] = block
    return resolved, ambiguous


def _incompatible(contributing) -> str | None:
    for name in PRECONDITIONS:
        values = {}
        for block in contributing:
            attr = block.attributes.get(name)
            if attr is not None:
                values.setdefault(norm(attr.value).casefold(), attr.value)
        if len(values) > 1:
            shown = " vs ".join(str(v) for v in values.values())
            return f"{name} differs across the sections summed ({shown})"
    return None


def _sum_losses(blocks, year: int, n, target_scale: float) -> float:
    """Σ Loss amount over every matching record. No record means 0, not absent."""
    total = 0.0
    for block in blocks:
        factor = scale_factor(block) / target_scale
        for record in match_records(block, str(year), n):
            value = record.values.get(LOSS_FIELD)
            if isinstance(value, (int, float)):
                total += float(value) * factor
    return total


def _sum_basis(blocks, year: int, n, target_scale: float) -> tuple[float | None, int]:
    """Σ Incurred Losses for one year, and how many sections actually reported it.

    One record per section per year — the bare year, not its ``9 months`` companion;
    ``match_record`` resolves that. Where a section has no row for the year, it does not
    contribute, and the count says so rather than the total quietly being short.
    """
    total, contributed = 0.0, 0
    for block in blocks:
        record = match_record(block, str(year), n)
        if record is None:
            continue
        value = record.values.get(BASIS_FIELD)
        if isinstance(value, (int, float)):
            total += float(value) * scale_factor(block) / target_scale
            contributed += 1
    return (total if contributed else None), contributed


def loss_share_tables(nomenclature, blocks) -> list[LossShareTable]:
    """One table per section group that has both a basis and at least one loss dataset."""
    if isinstance(blocks, dict):
        blocks = list(blocks.values())

    sections = getattr(nomenclature, "sections", None) or []
    if not sections:
        return []

    n = nomenclature.actual_year
    threshold = threshold_of(nomenclature)

    out = []
    for kind in dict.fromkeys(s.kind for s in sections):
        names = [s.name for s in sections if s.kind == kind]
        basis, ambiguous = _contributors(blocks, BASIS_ROLE, names)
        loss_blocks: dict[str, dict[str, Block]] = {}
        for role in LOSS_ROLES:
            resolved, more = _contributors(blocks, role, names)
            ambiguous.extend(more)
            if resolved:
                loss_blocks[role] = resolved

        if not basis or not loss_blocks:
            continue                       # nothing to compare — structure, not a gap

        roles = tuple(sorted(loss_blocks))
        contributing = list(basis.values()) + [
            b for by_section in loss_blocks.values() for b in by_section.values()
        ]
        sheets = tuple(dict.fromkeys(b.sheet_name for b in basis.values()))
        target_scale = scale_factor(next(iter(basis.values())))
        scale_attr = next(iter(basis.values())).attributes.get("Scale")

        table = LossShareTable(
            kind=kind,
            sections=tuple(sorted(set(basis) | {s for r in loss_blocks.values() for s in r})),
            roles=roles,
            sheets=sheets,
            threshold=threshold,
            scale=str(scale_attr.value) if scale_attr else "",
        )

        reason = "; ".join(ambiguous) or _incompatible(contributing)
        if reason:
            table.skipped = reason
            out.append(table)
            continue

        years = sorted({y for b in basis.values() for y in years_in(b)})
        for year in years:
            by_role = {
                role: _sum_losses(list(by_section.values()), year, n, target_scale)
                for role, by_section in loss_blocks.items()
            }
            incurred, contributed = _sum_basis(list(basis.values()), year, n, target_scale)
            table.rows.append(
                LossShareRow(year, by_role, sum(by_role.values()), incurred,
                             contributed, len(basis))
            )
        out.append(table)

    return out


def tables_for_sheet(tables, sheet_name: str) -> list[LossShareTable]:
    return [t for t in tables if sheet_name in t.sheets]
