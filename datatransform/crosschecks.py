"""Interdependencies between sheets. Specification_v1.md §10.1.

Rules are declared in ⟦RULES⟧ of sheet 00. Each side is a reference in one of three
forms — a field on a record, an attribute of a sheet, or a field summed over every
matching record — and ``{N}``, ``{N+1}``, ``{Y}`` resolve against the data.

Three things this module refuses to do:

* **Guess which record a pattern means.** Records match on leading number first, then
  suffix. Several candidates and no exact suffix is *ambiguous*, reported as such.
* **Compare incomparable figures.** ``Premium basis``, ``Loss basis``, ``Share basis``
  and ``Currency`` are a *precondition*. Scale is normalised; a differing basis skips
  the rule and says why.
* **Treat an absent dataset as a failure.** A rule whose dataset does not exist in this
  pack is *not applicable*, which is structure, not a problem.
"""

from __future__ import annotations

import re

from .model import (
    Block,
    Confidence,
    Hypothesis,
    Reference,
    Rule,
    RuleResult,
    parse_reference,
)
from .nomenclature import Nomenclature, norm

LEADING_NUMBER = re.compile(r"^\s*(\d+)\s*(.*)$")
PATTERN = re.compile(r"^\{N([+-]\d+)?\}\s*(.*)$")
YEAR_WILDCARD = re.compile(r"^\{Y\}\s*(.*)$")

RELATIONS = {
    "=": lambda a, b, tol: abs(a - b) <= tol,
    "==": lambda a, b, tol: abs(a - b) <= tol,
    ">=": lambda a, b, tol: a - b >= -tol,
    "<=": lambda a, b, tol: b - a >= -tol,
    ">": lambda a, b, tol: a - b > -tol,
    "<": lambda a, b, tol: b - a > -tol,
}

# Attributes that must agree before two figures may be compared.
PRECONDITIONS = ("Premium basis", "Loss basis", "Share basis", "Currency")

SCALES = {"1": 1.0, "1,000": 1_000.0, "1000": 1_000.0,
          "1,000,000": 1_000_000.0, "1000000": 1_000_000.0}


def split_label(label: str) -> tuple[int | None, str]:
    """``"2025 9 months"`` → ``(2025, "9 months")``."""
    match = LEADING_NUMBER.match(norm(label))
    if not match:
        return None, norm(label)
    return int(match.group(1)), match.group(2).strip()


def resolve_pattern(pattern: str, n: int) -> tuple[int, str]:
    """``"{N} re-est"`` with N=2025 → ``(2025, "re-est")``."""
    match = PATTERN.match(norm(pattern))
    if not match:
        raise ValueError(f"{pattern!r} does not start with {{N}} or {{N+1}}")
    return n + int(match.group(1) or 0), match.group(2).strip()


def is_wildcard(pattern: str) -> bool:
    return YEAR_WILDCARD.match(norm(pattern)) is not None


def substitute_year(pattern: str, year: int) -> str:
    """``"{Y} 9 months"`` with year 2023 → ``"{N} 9 months"`` anchored on 2023."""
    match = YEAR_WILDCARD.match(norm(pattern))
    suffix = match.group(1).strip() if match else ""
    return f"{year}{' ' + suffix if suffix else ''}"


def years_in(block: Block) -> list[int]:
    key = block.dataset.key_field
    seen = {
        split_label(r.values[key])[0]
        for r in block.records
        if isinstance(r.values.get(key), str)
    }
    return sorted(y for y in seen if y is not None)


def _candidates(block: Block, year: int, suffix: str):
    key = block.dataset.key_field
    out = []
    for record in block.records:
        label = record.values.get(key)
        if not isinstance(label, str):
            continue
        rec_year, rec_suffix = split_label(label)
        if rec_year != year:
            continue
        if suffix and rec_suffix.casefold() != suffix.casefold():
            continue
        out.append((record, rec_suffix))
    return out


def match_records(block: Block, pattern: str, n: int) -> list:
    """Every record matching the pattern — used by ``SUM``."""
    target = pattern if not PATTERN.match(norm(pattern)) else None
    if target is None:
        year, suffix = resolve_pattern(pattern, n)
    else:
        year, suffix = split_label(pattern)
        if year is None:
            return []
    return [r for r, _ in _candidates(block, year, suffix)]


def match_record(block: Block, pattern: str, n: int):
    """Exactly one record, or ``None`` — spec §9.2 S1, S10.

    A pattern with no suffix matches a lone record for that year whatever its own
    suffix, so both ``2026`` and ``2026 est`` resolve; where several records share
    the year, the suffix must be exact.
    """
    if PATTERN.match(norm(pattern)):
        year, suffix = resolve_pattern(pattern, n)
    else:
        year, suffix = split_label(pattern)
        if year is None:
            return None

    same_year = _candidates(block, year, "")
    if not same_year:
        return None
    if not suffix:
        if len(same_year) == 1:
            return same_year[0][0]
        exact = [r for r, s in same_year if not s]
        return exact[0] if len(exact) == 1 else None

    exact = [r for r, s in same_year if s.casefold() == suffix.casefold()]
    return exact[0] if len(exact) == 1 else None


def _match_detail(block: Block, pattern: str, n: int) -> str:
    if PATTERN.match(norm(pattern)):
        year, suffix = resolve_pattern(pattern, n)
    else:
        year, suffix = split_label(pattern)
    same_year = _candidates(block, year, "") if year is not None else []
    if not same_year:
        return f"no record for {year}"
    labels = ", ".join(s or "(no suffix)" for _, s in same_year)
    return f"{len(same_year)} records share {year} ({labels}) and none matches exactly"


def scale_factor(block: Block) -> float:
    attr = block.attributes.get("Scale")
    return SCALES.get(norm(attr.value), 1.0) if attr else 1.0


def _incompatible(left: Block, right: Block) -> str | None:
    for name in PRECONDITIONS:
        a, b = left.attributes.get(name), right.attributes.get(name)
        if a is None or b is None:
            continue
        if norm(a.value).casefold() != norm(b.value).casefold():
            return f"{name} differs ({a.value} vs {b.value})"
    return None


def _resolve_side(ref: Reference, block: Block, n: int):
    """(value, detail) — or (None, reason) when it cannot be resolved."""
    if ref.is_attribute:
        attr = block.attributes.get(ref.name)
        if attr is None:
            return None, f"{block.dataset.key} declares no {ref.name!r}"
        return norm(attr.value), str(attr.value)

    if ref.aggregate == "SUM":
        records = match_records(block, ref.record, n)
        if not records:
            return 0.0, f"no records for {ref.record} — treated as 0"
        total = sum(r.values[ref.name] for r in records
                    if isinstance(r.values.get(ref.name), (int, float)))
        return float(total), f"Σ {len(records)} record(s) = {total:,.0f}"

    record = match_record(block, ref.record, n)
    if record is None:
        return None, _match_detail(block, ref.record, n)
    value = record.values.get(ref.name)
    if value is None:
        return None, f"{ref.name} is empty on that record"
    return value, f"{value:,.0f}" if isinstance(value, (int, float)) else str(value)


def evaluate_one(rule: Rule, left_ref: Reference, right_ref: Reference,
                 blocks: dict[str, Block], n: int | None, year: int | None = None
                 ) -> RuleResult:
    left_block, right_block = blocks.get(left_ref.dataset_key), blocks.get(right_ref.dataset_key)
    absent = [r.dataset_key for r, b in ((left_ref, left_block), (right_ref, right_block))
              if b is None]
    if absent:
        return RuleResult(rule, "not applicable",
                          f"{', '.join(absent)} is not in this pack", year=year)

    left_value, left_detail = _resolve_side(left_ref, left_block, n)
    right_value, right_detail = _resolve_side(right_ref, right_block, n)
    for side, value, detail in (("left", left_value, left_detail),
                                ("right", right_value, right_detail)):
        if value is None:
            return RuleResult(rule, "skipped", f"{side}: {detail}", year=year)

    # Attribute comparison — a string equality, no scale or tolerance involved.
    if left_ref.is_attribute or right_ref.is_attribute:
        if not (left_ref.is_attribute and right_ref.is_attribute):
            return RuleResult(rule, "skipped",
                              "one side names an attribute and the other a figure", year=year)
        passed = str(left_value).casefold() == str(right_value).casefold()
        return RuleResult(rule, "passed" if passed else "failed",
                          f"{left_detail} {rule.relation} {right_detail}", year=year)

    reason = _incompatible(left_block, right_block)
    if reason:
        return RuleResult(rule, "skipped",
                          f"not comparable: {reason} — normalising it is a judgment, "
                          "so the rule is not evaluated", year=year)

    right_in_left_scale = right_value * scale_factor(right_block) / scale_factor(left_block)
    test = RELATIONS.get(rule.relation)
    if test is None:
        return RuleResult(rule, "skipped", f"unknown relation {rule.relation!r}", year=year)

    passed = test(float(left_value), float(right_in_left_scale), rule.tolerance)
    detail = (f"{left_value:,.0f} {rule.relation} {right_in_left_scale:,.0f} "
              f"(tolerance {rule.tolerance:,.0f})")
    return RuleResult(rule, "passed" if passed else "failed", detail,
                      float(left_value), float(right_in_left_scale), year=year)


def evaluate(rule: Rule, blocks: dict[str, Block], n: int | None) -> list[RuleResult]:
    try:
        left_ref, right_ref = rule.left_ref, rule.right_ref
    except Exception as exc:                                     # noqa: BLE001
        return [RuleResult(rule, "skipped", str(exc))]

    wildcard = is_wildcard(left_ref.record or "") or is_wildcard(right_ref.record or "")
    if not wildcard:
        if n is None and any("{N" in (r.record or "") for r in (left_ref, right_ref)):
            return [RuleResult(rule, "skipped",
                               "⟦GLOBAL⟧ declares no 'Actual year', so {N} cannot resolve")]
        return [evaluate_one(rule, left_ref, right_ref, blocks, n)]

    # {Y} — expand once per year present on the wildcard side
    source = left_ref if is_wildcard(left_ref.record or "") else right_ref
    block = blocks.get(source.dataset_key)
    if block is None:
        return [RuleResult(rule, "not applicable",
                           f"{source.dataset_key} is not in this pack")]

    results = []
    for year in years_in(block):
        left = left_ref.with_record(substitute_year(left_ref.record, year)) \
            if is_wildcard(left_ref.record or "") else left_ref
        right = right_ref.with_record(substitute_year(right_ref.record, year)) \
            if is_wildcard(right_ref.record or "") else right_ref
        results.append(evaluate_one(rule, left, right, blocks, n, year))
    return results


def run_rules(nomenclature: Nomenclature, blocks: dict[str, Block]) -> list[RuleResult]:
    n = nomenclature.actual_year
    out = []
    for rule in nomenclature.rules:
        out.extend(evaluate(rule, blocks, n))
    return out


def results_for(results, dataset_key: str) -> list[RuleResult]:
    """Every rule the dataset takes part in, either side."""
    out = []
    for result in results:
        for ref in (result.rule.left, result.rule.right):
            try:
                if parse_reference(ref).dataset_key == dataset_key:
                    out.append(result)
                    break
            except Exception:                                    # noqa: BLE001
                continue
    return out


def hypotheses_from(results, dataset_key: str) -> list[Hypothesis]:
    """A failed or unexpectedly skipped rule is a question — spec §10.1.

    *Not applicable* raises nothing: a rule that does not apply to this pack is
    structure, not a finding.
    """
    out = []
    for result in results_for(results, dataset_key):
        if result.status in ("passed", "not applicable"):
            continue
        out.append(
            Hypothesis(
                id="",
                dataset_key=dataset_key,
                attribute=f"Crosscheck {result.label}",
                value=result.rule.note or f"{result.rule.left} {result.rule.relation} "
                                          f"{result.rule.right}",
                confidence=Confidence.OPEN if result.status == "failed"
                else Confidence.ASSUMED,
                source="tool",
                note=f"{result.status}: {result.detail}",
            )
        )
    return out
