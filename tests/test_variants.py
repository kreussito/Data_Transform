"""The four treaty shapes — specification_v1.md §2.3.

The two Nat Cat workbooks carry the same sections and the same figures; they differ
only in whether the cat blocks share a sheet or sit in three. Everything downstream
must be identical, which is the test that a section really is just a block.
"""

from __future__ import annotations

import shutil
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

from datatransform.crosschecks import find_block, run_rules
from datatransform.extract import extract_sheet, last_non_empty_row
from datatransform.markers import read_markers
from datatransform.model import ExtractionError
from datatransform.nomenclature import read_nomenclature
from datatransform.runner import run
from datatransform.specs import step2_for
from datatransform.transform import apply_step2

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
    assert ("04", "Windstorm", 5) in shape(combined)   # 3 events, 5 underwriting years


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


def _written_blocks(path):
    wb = load_workbook(path, data_only=True)
    return {
        name: [c.value for row in wb[name].iter_rows() for c in row
               if isinstance(c.value, str) and c.value.startswith("Source: ")]
        for name in wb.sheetnames
    }


def test_every_section_block_of_a_stacked_sheet_is_written(tmp_path):
    """Two sections in one sheet means two outputs — the first must not be erased."""
    source = tmp_path / FIRE_CAT.name
    shutil.copy2(FIRE_CAT, source)
    out = tmp_path / "out.xlsx"
    run(source, out, tmp_path / "logs")

    written = _written_blocks(out)
    for sheet in ("01. History Cat", "02. EPI Cat", "04. Cat Losses"):
        assert len(written[sheet]) == 2, f"{sheet}: {written[sheet]}"
    assert len(written["01. History Fire"]) == 1


def test_rerunning_replaces_the_output_rather_than_stacking_it(tmp_path):
    source = tmp_path / FIRE_CAT.name
    shutil.copy2(FIRE_CAT, source)
    out = tmp_path / "out.xlsx"

    run(source, out, tmp_path / "logs")
    first = _written_blocks(out)
    run(out, tmp_path / "again.xlsx", tmp_path / "logs")   # output of a run, re-run
    assert _written_blocks(tmp_path / "again.xlsx") == first


def test_cat_losses_get_an_annual_table(tmp_path):
    source = tmp_path / FIRE_EQ_WIND.name
    shutil.copy2(FIRE_EQ_WIND, source)
    out = tmp_path / "out.xlsx"
    run(source, out, tmp_path / "logs")

    ws = load_workbook(out, data_only=True)["04. Cat Losses Wind"]
    text = [c.value for row in ws.iter_rows() for c in row if isinstance(c.value, str)]
    assert "Annual sum of cat losses" in text
    # Windstorm has events in 2021, 2022, 2024 and 2025; the quiet year still shows
    assert any("shown as 0" in t for t in text)


def test_the_process_log_documents_all_three_cat_tables(tmp_path):
    """The reviewer's record of what was done — spec §11."""
    source = tmp_path / FIRE_EQ_WIND.name
    shutil.copy2(FIRE_EQ_WIND, source)
    logs = tmp_path / "logs"
    run(source, tmp_path / "out.xlsx", logs)

    text = next(logs.glob("*_process.log")).read_text(encoding="utf-8")
    for title in ("Annual sum of cat losses", "By occurrence year", "By event"):
        assert f"aggregated: {title}" in text
    # The absent optional field is stated, not silently dropped.
    assert "Declared optional in sheet 00 and absent here: Number of Claims" in text


# ──────────────────────────────────── 04, events across underwriting years

def _cat_step2(path, section):
    nomenclature, blocks = _blocks(path)
    block = next(b for b in blocks if b.dataset.role == "04" and b.section == section)
    return block, apply_step2(block, step2_for(block.dataset.key), nomenclature, blocks)


def _table(result, title):
    return next(t for t in result.aggregates if t.title == title)


@pytest.mark.parametrize("path", [FIRE_CAT, FIRE_EQ_WIND])
def test_an_event_spanning_underwriting_years_is_reported_once_per_year(path):
    """Risks written in two years were both on cover, so the cedent reports twice."""
    _, result = _cat_step2(path, "Windstorm")
    bettina = [r for r in result.records if r.values["Event ID"] == "EV-202"]
    assert [(r.values["Year"], r.values["Loss amount"]) for r in bettina] == [
        ("2024", 1000.0), ("2025", 6600.0),
    ]
    # One event, one pair of dates — the storm did not happen twice.
    assert {(r.values["Event Date"], r.values["Event End Date"]) for r in bettina} == {
        (date(2024, 12, 30), date(2025, 1, 2)),
    }


@pytest.mark.parametrize("path", [FIRE_CAT, FIRE_EQ_WIND])
@pytest.mark.parametrize("section", ["Earthquake", "Windstorm"])
def test_the_three_cat_tables_answer_three_questions_from_one_total(path, section):
    """Underwriting year, occurrence year and event group the same losses differently."""
    _, result = _cat_step2(path, section)
    total = result.totals()["Loss amount"]
    assert [t.title for t in result.aggregates] == [
        "Annual sum of cat losses", "By occurrence year", "By event",
    ]
    for table in result.aggregates:
        assert table.totals()["Loss amount"] == total


@pytest.mark.parametrize("path", [FIRE_CAT, FIRE_EQ_WIND])
def test_the_by_event_table_shows_what_the_event_cost(path):
    """No annual row shows 7,600 — and that is the figure a cat layer is priced on."""
    _, result = _cat_step2(path, "Windstorm")
    by_event = dict(_table(result, "By event").rows)
    assert by_event["EV-202"]["Loss amount"] == 7600.0

    annual = dict(_table(result, "Annual sum of cat losses").rows)
    assert 7600.0 not in [v["Loss amount"] for v in annual.values()]
    assert annual["2024"]["Loss amount"] == 2800.0    # Bettina's share plus Kyrill II


@pytest.mark.parametrize("path", [FIRE_CAT, FIRE_EQ_WIND])
def test_occurrence_year_differs_from_underwriting_year(path):
    """The New Year straddle: written across two years, occurring in one."""
    _, result = _cat_step2(path, "Windstorm")
    by_occurrence = dict(_table(result, "By occurrence year").rows)
    assert by_occurrence["2024"]["Loss amount"] == 9400.0    # Bettina whole, plus Kyrill
    assert "2025" not in by_occurrence


def test_the_occurrence_year_attribute_redirects_the_grouping():
    """A cedent who counts an event by when it ended says so in column A — S15."""
    from datatransform.model import Attribute

    nomenclature, blocks = _blocks(FIRE_EQ_WIND)
    block = next(b for b in blocks if b.dataset.role == "04" and b.section == "Windstorm")
    block.attributes["Occurrence year from"] = Attribute(
        "Occurrence year from", "Event End Date", False, 0
    )
    result = apply_step2(block, step2_for(block.dataset.key), nomenclature, blocks)

    table = _table(result, "By occurrence year")
    assert table.group_by == "year(Event End Date)"
    assert dict(table.rows)["2025"]["Loss amount"] == 7600.0     # Bettina ended in 2025
    assert "2024" in dict(table.rows) and dict(table.rows)["2024"]["Loss amount"] == 1800.0


def test_an_attribute_naming_a_field_the_block_lacks_is_fatal():
    from datatransform.model import Attribute

    nomenclature, blocks = _blocks(FIRE_EQ_WIND)
    block = next(b for b in blocks if b.dataset.role == "04" and b.section == "Windstorm")
    block.attributes["Occurrence year from"] = Attribute(
        "Occurrence year from", "Date of Loss", False, 0
    )
    with pytest.raises(ExtractionError, match="does not extract"):
        apply_step2(block, step2_for(block.dataset.key), nomenclature, blocks)


def test_only_the_occurrence_date_is_checked_against_the_year():
    """An event ending in the next year is not misfiled — spec §10.2, S15."""
    from datatransform.extract import _check_occurrence_year
    from datatransform.model import Attribute

    _, blocks = _blocks(FIRE_EQ_WIND)
    block = next(b for b in blocks if b.dataset.role == "04" and b.section == "Windstorm")
    block.hypotheses.clear()
    block.attributes["Year basis"] = Attribute("Year basis", "Occurrence", False, 0)
    _check_occurrence_year(block)

    raised = [h for h in block.hypotheses if h.attribute == "Occurrence year mismatch"]
    assert len(raised) == 1
    # Eunice booked to 2021 and Bettina to 2025 disagree with their start dates; the
    # 2024 Bettina record does not, even though the storm ended in 2025.
    assert raised[0].value == "2 record(s)"
    assert "2025-01-02" not in raised[0].note


def test_the_occurrence_date_field_must_be_a_date_field_of_the_block():
    from datatransform.extract import _occurrence_date_field
    from datatransform.model import Attribute

    _, blocks = _blocks(FIRE_EQ_WIND)
    block = next(b for b in blocks if b.dataset.role == "04" and b.section == "Windstorm")
    assert _occurrence_date_field(block) == "Event Date"

    block.attributes["Occurrence year from"] = Attribute(
        "Occurrence year from", "Event Name", False, 0
    )
    with pytest.raises(ExtractionError, match="not a date field"):
        _occurrence_date_field(block)


# ───────────────────────────────────────────────────── optional fields

def test_an_absent_optional_field_is_not_a_gap():
    """Earthquake reports claim counts, Windstorm does not — spec §8 F4."""
    for path in (FIRE_CAT, FIRE_EQ_WIND):
        _, blocks = _blocks(path)
        by_section = {b.section: b for b in blocks if b.dataset.role == "04"}
        assert "Number of Claims" in by_section["Earthquake"].fields
        assert "Number of Claims" not in by_section["Windstorm"].fields
        assert not by_section["Windstorm"].unextracted
        claims = {r.values["Year"]: r.values["Number of Claims"]
                  for r in by_section["Earthquake"].records}
        assert claims == {"2021": 31.0, "2022": 12.0, "2023": 188.0}


def test_an_optional_field_is_declared_with_its_annotation_stripped():
    nomenclature, _ = _blocks(FIRE_CAT)
    dataset = nomenclature.datasets["04. Cat Losses"]
    assert dataset.headers[-1] == "Number of Claims"
    assert dataset.optional == ("Number of Claims",)
    assert dataset.is_optional("Number of Claims")
    assert not dataset.is_optional("Loss amount")


# ────────────────────────────────────────────────── R-10, the cat counterpart

def test_year_basis_must_agree_between_history_and_cat_losses():
    nomenclature, blocks = _blocks(FIRE_CAT)
    results = [r for r in run_rules(nomenclature, blocks) if r.rule.id == "R-10"]
    assert [r.section for r in results] == ["Earthquake", "Windstorm"]
    assert all(r.status == "passed" for r in results)


def test_the_engineering_pack_needs_no_cat_machinery(tmp_path):
    source = tmp_path / ENGINEERING.name
    shutil.copy2(ENGINEERING, source)
    report = run(source, tmp_path / "out.xlsx", tmp_path / "logs")

    processed = {o.sheet for o in report.outcomes if o.status == "processed"}
    assert processed == {"01. History", "02. EPI Projections", "03. Large Losses"}
    assert all(r.status != "failed" for r in report.rule_results)
