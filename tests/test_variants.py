"""The four treaty shapes — specification_v1.md §2.3.

The two Nat Cat workbooks carry the same sections and the same figures; they differ
only in whether the cat blocks share a sheet or sit in three. Everything downstream
must be identical, which is the test that a section really is just a block.
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest
from openpyxl import load_workbook

from datatransform.crosschecks import find_block, run_rules
from datatransform.extract import extract_sheet, last_non_empty_row
from datatransform.markers import read_markers
from datatransform.nomenclature import read_nomenclature
from datatransform.runner import run

ROOT = Path(__file__).resolve().parents[1]

FIRE = ROOT / "Intake_v1.xlsx"
ENGINEERING = ROOT / "Intake_Engineering_v1.xlsx"
FIRE_CAT = ROOT / "Intake_FireCat_v1.xlsx"
FIRE_EQ_WIND = ROOT / "Intake_FireEQWind_v1.xlsx"

ALL = [FIRE, ENGINEERING, FIRE_CAT, FIRE_EQ_WIND]


def _blocks(path):
    values = load_workbook(path, data_only=True)
    formulas = load_workbook(path, data_only=False)
    nomenclature = read_nomenclature(formulas)
    out = []
    for title in formulas.sheetnames:
        if title.startswith("00."):
            continue
        markers = read_markers(formulas[title], last_non_empty_row(values[title]))
        out.extend(extract_sheet(values[title], formulas[title], nomenclature, markers))
    return nomenclature, out


# ───────────────────────────────────────────────────────── ⟦SECTIONS⟧

@pytest.mark.parametrize("path,expected", [
    (FIRE, [("Fire", "per risk")]),
    (ENGINEERING, [("Engineering", "per risk")]),
    (FIRE_CAT, [("Fire", "per risk"), ("Earthquake", "cat"), ("Windstorm", "cat")]),
    (FIRE_EQ_WIND, [("Fire", "per risk"), ("Earthquake", "cat"), ("Windstorm", "cat")]),
])
def test_sections_are_declared_in_sheet_00(path, expected):
    nomenclature, _ = _blocks(path)
    assert [(s.name, s.kind) for s in nomenclature.sections] == expected


def test_a_section_declares_which_datasets_it_expects():
    nomenclature, _ = _blocks(FIRE_CAT)
    by_name = {s.name: s for s in nomenclature.sections}
    assert by_name["Fire"].roles == ("01", "02", "03")       # large losses
    assert by_name["Earthquake"].roles == ("01", "02", "04")  # event losses instead
    assert by_name["Fire"].expects("03") and not by_name["Earthquake"].expects("03")


@pytest.mark.parametrize("path", ALL)
def test_every_block_declares_its_section(path):
    nomenclature, blocks = _blocks(path)
    names = {s.name for s in nomenclature.sections}
    assert blocks
    for block in blocks:
        assert block.section in names, f"{block.dataset.key} declares {block.section!r}"


# ─────────────────────────────────────────── one sheet or three, same blocks

def test_the_two_cat_shapes_produce_the_same_section_blocks():
    """One combined Cat sheet, or three separate sheets — the same sections."""
    _, combined = _blocks(FIRE_CAT)
    _, separate = _blocks(FIRE_EQ_WIND)

    def shape(blocks):
        return sorted((b.dataset.role, b.section, len(b.records)) for b in blocks)

    assert shape(combined) == shape(separate)
    assert ("01", "Earthquake", 6) in shape(combined)
    assert ("04", "Windstorm", 3) in shape(combined)


def test_the_two_cat_shapes_carry_identical_figures():
    _, combined = _blocks(FIRE_CAT)
    _, separate = _blocks(FIRE_EQ_WIND)

    def values(blocks):
        return {
            (b.dataset.role, b.section): [r.values for r in b.records]
            for b in blocks
        }

    assert values(combined) == values(separate)


def test_stacked_blocks_do_not_swallow_each_other():
    """Two sections in one sheet must not share records — spec §6."""
    _, blocks = _blocks(FIRE_CAT)
    cat_history = [b for b in blocks if b.sheet_name == "01. History Cat"]
    assert len(cat_history) == 2
    assert {b.section for b in cat_history} == {"Earthquake", "Windstorm"}
    assert all(len(b.records) == 6 for b in cat_history)
    first, second = sorted(cat_history, key=lambda b: b.index)
    assert first.records[0].values["Premium"] != second.records[0].values["Premium"]


# ─────────────────────────────────────────────────── role resolution

def test_a_role_resolves_through_its_section():
    _, blocks = _blocks(FIRE_EQ_WIND)
    for section, key in [("Fire", "01 History Fire"),
                         ("Earthquake", "01 History EQ"),
                         ("Windstorm", "01 History Wind")]:
        assert find_block(blocks, "01", section).dataset.key == key


def test_a_role_without_a_section_is_ambiguous_across_sections():
    _, blocks = _blocks(FIRE_EQ_WIND)
    assert find_block(blocks, "01", None) is None      # three candidates, no guess


def test_a_full_key_still_resolves():
    _, blocks = _blocks(FIRE_CAT)
    assert find_block(blocks, "03 Large Fire", "Fire").dataset.key == "03 Large Fire"


# ───────────────────────────────────────────── rule scoping and expansion

def test_rules_expand_once_per_applicable_section():
    nomenclature, blocks = _blocks(FIRE_EQ_WIND)
    results = run_rules(nomenclature, blocks)
    by_rule = {}
    for r in results:
        by_rule.setdefault(r.rule.id, []).append(r.section)

    assert by_rule["R-05"] == ["Fire", "Earthquake", "Windstorm"]   # scope: all
    assert set(by_rule["R-01"]) == {"Fire"}                         # scope: per risk
    assert set(by_rule["R-02"]) == {"Earthquake", "Windstorm"}      # scope: cat


def test_a_cat_rule_is_not_applicable_on_a_per_risk_treaty():
    for path in (FIRE, ENGINEERING):
        nomenclature, blocks = _blocks(path)
        results = run_rules(nomenclature, blocks)
        cat = [r for r in results if r.rule.id == "R-02"]
        assert [r.status for r in cat] == ["not applicable"], path.name
        assert "matches no section" in cat[0].detail


def test_a_per_risk_rule_never_runs_on_a_cat_section():
    nomenclature, blocks = _blocks(FIRE_CAT)
    results = run_rules(nomenclature, blocks)
    assert all(r.section == "Fire" for r in results if r.rule.id in ("R-01", "R-09"))


@pytest.mark.parametrize("path", ALL)
def test_every_applicable_rule_passes(path):
    nomenclature, blocks = _blocks(path)
    results = run_rules(nomenclature, blocks)
    failed = [(r.label, r.detail) for r in results if r.status == "failed"]
    skipped = [(r.label, r.detail) for r in results if r.status == "skipped"]
    assert not failed, f"{path.name}: {failed}"
    assert not skipped, f"{path.name}: {skipped}"


def test_the_two_cat_shapes_produce_identical_crosschecks():
    outcome = []
    for path in (FIRE_CAT, FIRE_EQ_WIND):
        nomenclature, blocks = _blocks(path)
        outcome.append(sorted(
            (r.rule.id, r.section, r.year, r.status, r.left_value, r.right_value)
            for r in run_rules(nomenclature, blocks)
        ))
    assert outcome[0] == outcome[1]


# ─────────────────────────────────────────────────────── end to end

@pytest.mark.parametrize("path", ALL)
def test_each_treaty_shape_runs_end_to_end(path, tmp_path):
    source = tmp_path / path.name
    shutil.copy2(path, source)
    report = run(source, tmp_path / "out.xlsx", tmp_path / "logs")

    assert report.ok, [o.detail for o in report.outcomes if o.status == "error"]
    assert report.rules_ok
    assert any(o.status == "processed" for o in report.outcomes)


def test_cat_losses_get_an_annual_table(tmp_path):
    source = tmp_path / FIRE_EQ_WIND.name
    shutil.copy2(FIRE_EQ_WIND, source)
    out = tmp_path / "out.xlsx"
    run(source, out, tmp_path / "logs")

    ws = load_workbook(out, data_only=True)["04. Cat Losses Wind"]
    text = [c.value for row in ws.iter_rows() for c in row if isinstance(c.value, str)]
    assert "Annual sum of cat losses" in text
    # Windstorm has events in 2022, 2024 and 2025; the quiet years still show
    assert any("shown as 0" in t for t in text)


def test_the_engineering_pack_needs_no_cat_machinery(tmp_path):
    source = tmp_path / ENGINEERING.name
    shutil.copy2(ENGINEERING, source)
    report = run(source, tmp_path / "out.xlsx", tmp_path / "logs")

    processed = {o.sheet for o in report.outcomes if o.status == "processed"}
    assert processed == {"01. History", "02. EPI Projections", "03. Large Losses"}
    assert all(r.status != "failed" for r in report.rule_results)
