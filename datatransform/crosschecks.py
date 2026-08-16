"""Interdependencies between sheets. Specification_v1.md §10.1.

Rules are declared in ⟦RULES⟧ of sheet 00 as
``<dataset key>.<field>@<record>  <rel>  <dataset key>.<field>@<record>``, with
``{N}`` and ``{N+1}`` resolving from ⟦GLOBAL⟧'s ``Actual year``.

Two things this module refuses to do:

* **Guess which record a pattern means.** Records match on leading number first, then
  suffix — the same leading-number rule S10 uses.
* **Compare incomparable figures.** ``Premium basis``, ``Currency`` and ``Scale`` are a
  *precondition*, not part of the comparison. Scale is normalised; a differing basis or
  currency skips the rule and says why. Quietly comparing GWP against GNPI would be
  worse than not checking at all.
"""

from __future__ import annotations

import re

from .model import Block, Confidence, Hypothesis, Rule, RuleResult
from .nomenclature import Nomenclature, norm

LEADING_NUMBER = re.compile(r"^\s*(\d+)\s*(.*)$")
PATTERN = re.compile(r"^\{N([+-]\d+)?\}\s*(.*)$")

RELATIONS = {
    "=": lambda a, b, tol: abs(a - b) <= tol,
    "==": lambda a, b, tol: abs(a - b) <= tol,
    ">=": lambda a, b, tol: a - b >= -tol,
    "<=": lambda a, b, tol: b - a >= -tol,
    ">": lambda a, b, tol: a - b > -tol,
    "<": lambda a, b, tol: b - a > -tol,
}

# Attributes that must agree before two figures may be compared.
PRECONDITIONS = ("Premium basis", "Loss basis", "Currency")

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
    offset = int(match.group(1) or 0)
    return n + offset, match.group(2).strip()


def match_record(block: Block, pattern: str, n: int):
    """Leading number first, suffix second — spec §9.2 S1, S10.

    A pattern with no suffix matches a lone record for that year whatever its own
    suffix, so both ``2026`` and ``2026 est`` resolve; where several records share the
    year, the suffix must match exactly.
    """
    year, suffix = resolve_pattern(pattern, n)
    key = block.dataset.key_field

    same_year = []
    for record in block.records:
        label = record.values.get(key)
        if not isinstance(label, str):
            continue
        rec_year, rec_suffix = split_label(label)
        if rec_year == year:
            same_year.append((record, rec_suffix))

    if not same_year:
        return None
    if not suffix and len(same_year) == 1:
        return same_year[0][0]

    exact = [r for r, s in same_year if s.casefold() == suffix.casefold()]
    return exact[0] if len(exact) == 1 else None


def scale_factor(block: Block) -> float:
    attr = block.attributes.get("Scale")
    if attr is None:
        return 1.0
    return SCALES.get(norm(attr.value), 1.0)


def _incompatible(left: Block, right: Block) -> str | None:
    for name in PRECONDITIONS:
        a, b = left.attributes.get(name), right.attributes.get(name)
        if a is None or b is None:
            continue
        if norm(a.value).casefold() != norm(b.value).casefold():
            return f"{name} differs ({a.value} vs {b.value})"
    return None


def evaluate(rule: Rule, blocks: dict[str, Block], n: int | None) -> RuleResult:
    if n is None:
        return RuleResult(rule, "skipped",
                          "⟦GLOBAL⟧ declares no 'Actual year', so {N} cannot resolve")

    try:
        left_key, left_field, left_rec = Rule.parse_ref(rule.left)
        right_key, right_field, right_rec = Rule.parse_ref(rule.right)
    except Exception as exc:                                     # noqa: BLE001
        return RuleResult(rule, "skipped", str(exc))

    left_block, right_block = blocks.get(left_key), blocks.get(right_key)
    missing = [k for k, b in ((left_key, left_block), (right_key, right_block)) if b is None]
    if missing:
        return RuleResult(rule, "skipped", f"dataset {', '.join(missing)} not extracted")

    reason = _incompatible(left_block, right_block)
    if reason:
        return RuleResult(rule, "skipped",
                          f"not comparable: {reason} — normalising it is a judgment, "
                          "so the rule is not evaluated")

    left_record = match_record(left_block, left_rec, n)
    right_record = match_record(right_block, right_rec, n)
    for label, record, pattern in (("left", left_record, rule.left),
                                   ("right", right_record, rule.right)):
        if record is None:
            return RuleResult(rule, "skipped", f"no {label} record matching {pattern!r}")

    left_value = left_record.values.get(left_field)
    right_value = right_record.values.get(right_field)
    if not isinstance(left_value, (int, float)) or not isinstance(right_value, (int, float)):
        return RuleResult(rule, "skipped",
                          f"{left_field} or {right_field} is not numeric on these records")

    # Compare in the left block's scale.
    left_scale, right_scale = scale_factor(left_block), scale_factor(right_block)
    right_in_left_scale = right_value * right_scale / left_scale

    test = RELATIONS.get(rule.relation)
    if test is None:
        return RuleResult(rule, "skipped", f"unknown relation {rule.relation!r}")

    passed = test(float(left_value), float(right_in_left_scale), rule.tolerance)
    detail = (f"{left_value:,.0f} {rule.relation} {right_in_left_scale:,.0f} "
              f"(tolerance {rule.tolerance:,.0f})")
    return RuleResult(rule, "passed" if passed else "failed", detail,
                      float(left_value), float(right_in_left_scale))


def run_rules(nomenclature: Nomenclature, blocks: dict[str, Block]) -> list[RuleResult]:
    n = nomenclature.actual_year
    return [evaluate(rule, blocks, n) for rule in nomenclature.rules]


def results_for(results, dataset_key: str) -> list[RuleResult]:
    """Every rule the dataset takes part in, either side."""
    out = []
    for result in results:
        for ref in (result.rule.left, result.rule.right):
            try:
                key, _, _ = Rule.parse_ref(ref)
            except Exception:                                    # noqa: BLE001
                continue
            if key == dataset_key:
                out.append(result)
                break
    return out


def hypotheses_from(results, dataset_key: str) -> list[Hypothesis]:
    """A skipped or failed rule is a question, not a silent pass — spec §10.1."""
    out = []
    for result in results_for(results, dataset_key):
        if result.status == "passed":
            continue
        out.append(
            Hypothesis(
                id="",
                dataset_key=dataset_key,
                attribute=f"Crosscheck {result.rule.id}",
                value=result.rule.note or f"{result.rule.left} {result.rule.relation} "
                                          f"{result.rule.right}",
                confidence=Confidence.OPEN if result.status == "failed"
                else Confidence.ASSUMED,
                source="tool",
                note=f"{result.status}: {result.detail}",
            )
        )
    return out
