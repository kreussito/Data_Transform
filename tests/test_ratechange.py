"""09. Rate Development — specification_v1.md §2.7, §10.5.

The first dataset that is not the cedent's. The underwriter types it in, from a quote, a
broker note or their own estimate, so where each figure came from is itself recorded —
and it is checked against the one number a submission cannot fake, the rate change that
premium and exposure between them imply.
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest
from openpyxl import load_workbook

from datatransform.constants import RATE_CLAIM_DEFAULT, RATE_CLAIM_WARNING
from datatransform.extract import extract_sheet, last_non_empty_row
from datatransform.growth import growth_tables
from datatransform.markers import read_markers
from datatransform.model import Attribute, ExtractionError
from datatransform.nomenclature import read_nomenclature
from datatransform.ratechange import rate_claims, scope_of, threshold_of
from datatransform.runner import run
from datatransform.specs import step2_for
from datatransform.transform import apply_step2

ROOT = Path(__file__).resolve().parents[1]
FULL = ROOT / "Intake_FireCatFull_v1.xlsx"


def _read(path=FULL):
    values = load_workbook(path, data_only=True)
    formulas = load_workbook(path, data_only=False)
    nomenclature = read_nomenclature(formulas)
    blocks = []
    for title in formulas.sheetnames:
        if title.startswith("00."):
            continue
        markers = read_markers(formulas[title], last_non_empty_row(values[title]))
        blocks.extend(extract_sheet(values[title], formulas[title], nomenclature, markers))
    return nomenclature, blocks


def _rate_blocks(blocks):
    return sorted((b for b in blocks if b.dataset.role == "09"), key=lambda b: b.index)


def _claims(nomenclature=None, blocks=None):
    if nomenclature is None:
        nomenclature, blocks = _read()
    results = [apply_step2(b, step2_for(b.dataset.key), nomenclature, blocks)
               for b in blocks if step2_for(b.dataset.key)]
    growth = growth_tables(nomenclature, blocks, results)
    return {" + ".join(c.scope) or "?": c
            for c in rate_claims(nomenclature, blocks, growth)}


# ───────────────────────────── the two shapes a rate sheet arrives in

def test_a_block_may_carry_the_rate_or_the_change():
    _, blocks = _read()
    fire, cat = _rate_blocks(blocks)
    assert fire.fields == ("Year", "Rate")             # levels only
    assert cat.fields == ("Year", "Rate change")       # the movement only


def test_the_change_is_worked_out_from_the_levels():
    """1.2150 after 1.1800 is +2.97%, and the first year has nothing before it."""
    nomenclature, blocks = _read()
    fire = _rate_blocks(blocks)[0]
    result = apply_step2(fire, step2_for(fire.dataset.key), nomenclature, blocks)

    changes = [r.values.get("Rate change") for r in result.records]
    assert changes[0] is None
    assert changes[1] == pytest.approx(1.2150 / 1.1800 - 1, abs=1e-9)
    assert changes[-1] == pytest.approx(1.4060 / 1.3520 - 1, abs=1e-9)
    assert any("worked out from Rate" in n for n in result.notes)


def test_a_reported_change_is_never_overwritten():
    nomenclature, blocks = _read()
    cat = _rate_blocks(blocks)[1]
    result = apply_step2(cat, step2_for(cat.dataset.key), nomenclature, blocks)

    assert [r.values["Rate change"] for r in result.records] == [0.031, 0.042, 0.055, 0.040]
    assert any("as reported" in n for n in result.notes)
    assert "Rate change" not in result.derived_fields


def test_a_gap_in_the_years_is_named_rather_than_read_as_annual():
    nomenclature, blocks = _read()
    fire = _rate_blocks(blocks)[0]
    fire.records = [r for r in fire.records if r.values["Year"] != "2024"]

    result = apply_step2(fire, step2_for(fire.dataset.key), nomenclature, blocks)
    note = next(n for n in result.notes if "spans more than one year" in n)
    assert "2023→2025" in note and "must not be read as such" in note


def test_a_rate_is_a_number_but_not_a_quantity():
    """Five years' rates do not add up to a rate — §2.4's rule, applied to §2.7."""
    _, blocks = _read()
    fire, cat = _rate_blocks(blocks)
    assert fire.measure_fields == ()
    assert cat.measure_fields == ()
    assert fire.totals() == {}


# ──────────────────────────────────────── scope: one section or several

def test_a_scope_may_name_several_sections():
    nomenclature, blocks = _read()
    fire, cat = _rate_blocks(blocks)
    assert scope_of(fire, nomenclature) == ("Fire",)
    assert scope_of(cat, nomenclature) == ("Earthquake", "Windstorm")


def test_a_scope_naming_an_undeclared_section_is_fatal():
    """A rate compared against the wrong book is worse than no comparison."""
    nomenclature, blocks = _read()
    cat = _rate_blocks(blocks)[1]
    cat.attributes["Scope"] = Attribute("Scope", "Earthquake + Hurricane", False, 0)
    with pytest.raises(ExtractionError, match="⟦SECTIONS⟧ does not declare"):
        scope_of(cat, nomenclature)


def test_the_underwriter_source_is_recorded():
    nomenclature, blocks = _read()
    fire, cat = _rate_blocks(blocks)
    assert fire.attributes["Source"].value == "cedent"
    assert cat.attributes["Source"].value == "broker"
    assert _claims()["Earthquake + Windstorm"].source == "broker"


# ─────────────────────────── claimed against implied — spec §10.5

def test_the_claim_is_the_renewal_year():
    claim = _claims()["Earthquake + Windstorm"]
    assert claim.year == "2026"
    assert claim.claimed == pytest.approx(0.040)


def test_premium_adds_but_exposure_is_weighted():
    """A coastal risk sits in both aggregates, so their sum is not a portfolio figure —
    only the growth is used, in which a stable double count cancels."""
    claim = _claims()["Earthquake + Windstorm"]
    assert claim.combined
    assert {p[0] for p in claim.parts} == {"Earthquake", "Windstorm"}

    # The exposure growth is the mean of the two, weighted by each section's own first
    # version — never the growth of their sum, which would double count the coast.
    _, _, growths = zip(*claim.parts)
    assert min(growths) <= claim.exposure_growth <= max(growths)

    # 06 reports no Total of its own — it is derived in step 2, so the exposure has to
    # be read from there, exactly as the runner does it.
    nomenclature, blocks = _read()
    results = [apply_step2(b, step2_for(b.dataset.key), nomenclature, blocks)
               for b in blocks if step2_for(b.dataset.key)]
    tables = {g.section: g for g in growth_tables(nomenclature, blocks, results)}
    weights = {s: tables[s].versions[0].exposure for s in ("Earthquake", "Windstorm")}
    expected = sum(tables[s].exposure_growth * w for s, w in weights.items()) \
        / sum(weights.values())
    assert claim.exposure_growth == pytest.approx(expected)
    # Premium is the ratio of the two sums, not the mean of the two ratios.
    assert claim.premium_growth == pytest.approx((7900 + 9250) / (7250 + 8600) - 1)


def test_the_gap_is_measured_in_percentage_points():
    claim = _claims()["Earthquake + Windstorm"]
    assert claim.implied == pytest.approx(
        (1 + claim.premium_growth) / (1 + claim.exposure_growth) - 1)
    assert claim.gap == pytest.approx(claim.claimed - claim.implied)
    assert claim.status == "WARNING"          # +4.0% claimed against about −4%


def test_the_question_names_the_direction_and_the_likely_reasons():
    claim = _claims()["Earthquake + Windstorm"]
    assert "QUESTION FOR THE UNDERWRITER" in claim.question
    assert "above" in claim.question
    assert "risk-adjusted" in claim.question


def test_a_scope_with_no_exposure_is_shown_unchecked_not_dropped():
    """Fire has no cat aggregate, so §10.4 produced nothing to hold it against."""
    claim = _claims()["Fire"]
    assert claim.skipped
    assert "shown unchecked rather than dropped" in claim.skipped


def test_the_threshold_is_declared_in_global():
    nomenclature, _ = _read()
    assert threshold_of(nomenclature) == pytest.approx(0.05)
    assert RATE_CLAIM_WARNING == "Rate claim warning"
    assert RATE_CLAIM_DEFAULT == 0.05


def test_a_claim_within_the_threshold_raises_no_warning():
    nomenclature, blocks = _read()
    cat = _rate_blocks(blocks)[1]
    for record in cat.records:
        if record.values["Year"] == "2026":
            record.values["Rate change"] = -0.041      # about what the figures imply
    claim = _claims(nomenclature, blocks)["Earthquake + Windstorm"]
    assert claim.status == "within limits"
    assert abs(claim.gap) < 0.05


# ───────────────────────────────────────────────────────── end to end

def test_the_block_is_written_with_somewhere_to_answer(tmp_path):
    source = tmp_path / FULL.name
    shutil.copy2(FULL, source)
    out = tmp_path / "out.xlsx"
    report = run(source, out, tmp_path / "logs")

    assert report.ok, [o.detail for o in report.outcomes if o.status == "error"]
    assert len(report.rate_claims) == 2

    ws = load_workbook(out, data_only=True)["09. Rate Development"]
    text = [c.value for row in ws.iter_rows() for c in row if isinstance(c.value, str)]
    assert any("CLAIMED RATE CHANGE AGAINST IMPLIED — EARTHQUAKE + WINDSTORM" in t
               for t in text)
    assert any("QUESTION FOR THE UNDERWRITER" in t for t in text)
    assert any("never read as a portfolio figure" in t for t in text)
    assert "Answer:" in text


def test_the_process_log_records_both_scopes(tmp_path):
    source = tmp_path / FULL.name
    shutil.copy2(FULL, source)
    logs = tmp_path / "logs"
    run(source, tmp_path / "out.xlsx", logs)

    text = next(logs.glob("*_process.log")).read_text(encoding="utf-8")
    assert "Claimed rate change against implied (§10.5)" in text
    assert "Earthquake + Windstorm [broker]" in text
    assert "Fire — NOT EVALUATED" in text
