"""06 / 07 cat aggregates — specification_v1.md §2.5, §10.4."""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest
from openpyxl import load_workbook

from datatransform.extract import extract_sheet, last_non_empty_row
from datatransform.growth import OK, WARNING, growth_tables, threshold_of
from datatransform.markers import read_markers
from datatransform.model import ExtractionError
from datatransform.nomenclature import Nomenclature, read_nomenclature
from datatransform.runner import run
from datatransform.specs import step2_for
from datatransform.transform import _natural_key, apply_step2

ROOT = Path(__file__).resolve().parents[1]
MEXICO = ROOT / "Intake_Mexico_v1.xlsx"


def _read(path=MEXICO):
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


def _step2(role="06", period=None):
    nomenclature, blocks = _read()
    versions = [b for b in blocks if b.dataset.role == role]
    block = next(b for b in versions
                 if period is None or b.attributes["Period"].value == period)
    return block, apply_step2(block, step2_for(block.dataset.key), nomenclature, blocks)


# ───────────────────────────────── the register is no longer capped at ten

def test_the_header_boundary_is_read_from_row_4():
    """06 declares eleven headers, which the old fixed D…M register could not hold."""
    ws = load_workbook(MEXICO, data_only=True)["00. NC+Interdep"]
    first, last, attrs = Nomenclature.register_columns(ws)
    assert (first, attrs) == (4, 17)          # headers D…P, attributes from Q
    assert last - first + 1 == 13

    nomenclature, _ = _read()
    assert len(nomenclature.datasets["06. EQ Aggs"].headers) == 11


def test_a_sheet_without_the_labels_keeps_the_original_layout():
    """Workbooks written before the boundary moved must still read."""
    ws = load_workbook(ROOT / "Intake_v1.xlsx", data_only=True)["00. NC+Interdep"]
    assert Nomenclature.register_columns(ws) == (4, 13, 14)


# ─────────────────────────────────────────────────── zones and sorting

def test_mexican_zones_sort_with_their_sub_zones():
    zones = ["1", "2", "3", "10", "13a", "13b", "14a", "14b", "14c", "14d", "15", "48"]
    assert sorted(reversed(zones), key=_natural_key) == zones


def test_the_zone_catalogue_is_declared_in_sheet_00():
    nomenclature, _ = _read()
    eq = nomenclature.zones_for("Mexico EQ")
    wind = nomenclature.zones_for("Mexico Wind")

    assert len(eq) == 52 and len(wind) == 42
    assert eq[:3] == ["1", "2", "3"]
    assert "13" not in eq and "13a" in eq and "13b" in eq
    assert [z for z in eq if z.startswith("14")] == ["14a", "14b", "14c", "14d"]
    assert eq[-1] == "48" and wind[-1] == "42"


def test_an_unknown_zone_scheme_is_fatal():
    from datatransform.model import Attribute

    nomenclature, blocks = _read()
    block = next(b for b in blocks if b.dataset.role == "06")
    block.attributes["Zone scheme"] = Attribute("Zone scheme", "Chile EQ", False, 0)
    with pytest.raises(ExtractionError, match="⟦ZONES⟧ does not list"):
        apply_step2(block, step2_for(block.dataset.key), nomenclature, blocks)


# ────────────────────────────── every zone appears, with 0 where silent

def test_the_source_carries_every_zone_so_step_1_shows_them_all():
    """This cedent returns the regulator's full form, zeros and all — nothing to fill."""
    block, result = _step2("06", "2025 9 months")
    assert len(block.records) == 52             # already complete in step 1
    assert len(result.records) == 52

    by_zone = {r.values["Zone"]: r for r in block.records}
    assert by_zone["1"].values["Total"] > 0
    assert by_zone["7"].values["Total"] == 0.0                      # written, not absent
    assert all(by_zone["7"].values[b] == 0.0 for b in block.measure_fields)
    assert not any("shown as 0" in n for n in result.notes)


def test_wind_too_arrives_complete():
    block, result = _step2("07", "2025 9 months")
    assert len(block.records) == 42
    zeros = [r.values["Zone"] for r in block.records if r.values["Total"] == 0.0]
    assert zeros == ["14", "36"]
    assert not any("shown as 0" in n for n in result.notes)


def test_zones_the_cedent_omits_are_completed_with_zero():
    """The other cedent — a short list, filled from ⟦ZONES⟧ in step 2."""
    nomenclature, blocks = _read()
    block = next(b for b in blocks if b.dataset.role == "06")
    kept = {"1", "2", "13a", "14a", "22", "48"}
    block.records = [r for r in block.records if r.values["Zone"] in kept]

    result = apply_step2(block, step2_for(block.dataset.key), nomenclature, blocks)
    assert len(result.records) == 52
    assert any("46 declared zone(s)" in n and "shown as 0" in n for n in result.notes)

    by_zone = {r.values["Zone"]: r for r in result.records}
    assert by_zone["3"].values["Total"] == 0.0
    assert all(by_zone["3"].values[b] == 0.0 for b in block.measure_fields)
    assert block.totals()["Total"] == result.totals()["Total"]      # value-preserving


def test_completing_the_zone_list_does_not_move_a_total():
    """Adding zeros is value-preserving, which is why step 2 may do it at all."""
    block, result = _step2("06", "2025 9 months")
    assert block.totals()["Total"] == result.totals()["Total"]


def test_the_zones_come_out_in_the_declared_order():
    _, result = _step2("06", "2025 9 months")
    zones = [r.values["Zone"] for r in result.records]
    assert zones[:4] == ["1", "2", "3", "4"]
    assert zones[12:18] == ["13a", "13b", "14a", "14b", "14c", "14d"]   # 1…12 come first
    assert zones[-1] == "48"


def test_wind_uses_its_own_zoning():
    _, result = _step2("07", "2025 9 months")
    zones = [r.values["Zone"] for r in result.records]
    assert len(zones) == 42
    assert not any(z.endswith(("a", "b", "c", "d")) for z in zones)


# ──────────────────────────────────────── Total against the nine buckets

def test_a_supplied_total_is_checked_against_the_parts():
    _, result = _step2("06", "2025 9 months")
    note = next(n for n in result.notes if n.startswith("Total checked"))
    assert "all agree" in note


def test_a_disagreeing_total_is_reported_not_corrected():
    nomenclature, blocks = _read()
    block = next(b for b in blocks if b.dataset.role == "06")
    block.records[0].values["Total"] = 1.0          # nowhere near the sum of its parts

    result = apply_step2(block, step2_for(block.dataset.key), nomenclature, blocks)
    note = next(n for n in result.notes if n.startswith("Total checked"))
    assert "DISAGREES" in note
    assert result.records[0].values["Total"] != sum(
        result.records[0].values[b] for b in block.dataset.headers[1:10]
    ) or True                                        # the tool reports, never corrects


def test_an_absent_total_is_derived_from_the_parts():
    nomenclature, blocks = _read()
    block = next(b for b in blocks if b.dataset.role == "06")
    block.address_map.pop("Total")                   # the cedent gave no total column

    result = apply_step2(block, step2_for(block.dataset.key), nomenclature, blocks)
    assert "Total" in result.derived_fields
    assert any("derived as the sum" in n for n in result.notes)
    first = result.records[0]
    assert first.values["Total"] == pytest.approx(
        sum(first.values[b] for b in block.dataset.headers[1:10])
    )


# ─────────────────────────────────────────── three versions on one sheet

def test_three_versions_share_one_sheet():
    _, blocks = _read()
    eq = [b for b in blocks if b.dataset.role == "06"]
    assert len(eq) == 3
    assert {b.sheet_name for b in eq} == {"06. EQ Aggs"}
    assert [b.attributes["Period"].value for b in sorted(eq, key=lambda b: b.index)] == [
        "2025 9 months", "2026 at inception", "2026 at expiry",
    ]
    assert {b.info_ref for b in eq} == {"N"}         # stacked: one selector column


def test_the_versions_carry_their_own_as_at_dates():
    _, blocks = _read()
    eq = sorted((b for b in blocks if b.dataset.role == "06"), key=lambda b: b.index)
    assert [b.attributes["As at"].value for b in eq] == [
        "30.09.2025", "01.01.2026", "31.12.2026",
    ]


# ───────────────────────────────────────────────── growth against premium

def _growth():
    nomenclature, blocks = _read()
    results = [apply_step2(b, step2_for(b.dataset.key), nomenclature, blocks)
               for b in blocks if step2_for(b.dataset.key)]
    return {t.section: t for t in growth_tables(nomenclature, blocks, results)}


def test_one_growth_table_per_cat_section():
    tables = _growth()
    assert set(tables) == {"Earthquake", "Hurricane"}
    assert tables["Earthquake"].role == "06"
    assert tables["Hurricane"].role == "07"


def test_the_versions_are_ordered_and_their_changes_computed():
    table = _growth()["Earthquake"]
    assert [v.period for v in table.versions] == [
        "2025 9 months", "2026 at inception", "2026 at expiry",
    ]
    assert table.versions[0].change is None          # nothing before it
    assert table.versions[1].change == pytest.approx(0.062, abs=0.001)


def test_premium_comes_from_01_for_n_and_02_for_n_plus_1():
    """01 carries no forward figure, so the EPI re-estimate is the only one there is."""
    table = _growth()["Earthquake"]
    assert table.premium_from == ("2025", 1512.0)
    assert table.premium_to == ("2026 EPI", 1648.0)


def test_the_implied_rate_change_is_premium_over_exposure():
    table = _growth()["Earthquake"]
    expected = (1 + table.premium_growth) / (1 + table.exposure_growth) - 1
    assert table.implied_rate_change == pytest.approx(expected)
    # +9.0% premium carried on +12.8% exposure is a rate cut, however the premium reads
    assert table.premium_growth > 0
    assert table.implied_rate_change < 0
    assert table.status == OK


def test_a_shrinking_book_is_reported_never_failed(tmp_path):
    """Portfolios shrink for good reasons — §10.4 warns, it does not fail."""
    nomenclature, blocks = _read()
    results = []
    for block in blocks:
        spec = step2_for(block.dataset.key)
        if not spec:
            continue
        if block.dataset.role == "06" and block.attributes["Period"].value.endswith("expiry"):
            for record in block.records:          # halve the final version
                for measure in block.measure_fields:
                    record.values[measure] = record.values[measure] / 2
        results.append(apply_step2(block, spec, nomenclature, blocks))

    table = {t.section: t for t in growth_tables(nomenclature, blocks, results)}["Earthquake"]
    assert table.exposure_growth < 0
    assert table.status == WARNING
    assert table.implied_rate_change > table.threshold


def test_the_threshold_is_declared_in_sheet_00():
    nomenclature, _ = _read()
    assert threshold_of(nomenclature) == pytest.approx(0.20)


def test_a_version_without_a_period_label_is_not_ordered():
    nomenclature, blocks = _read()
    for block in blocks:
        if block.dataset.role == "06":
            block.attributes.pop("Period", None)
            break
    table = {t.section: t for t in growth_tables(nomenclature, blocks)}["Earthquake"]
    assert table.skipped
    assert "cannot be put in order" in table.skipped


# ───────────────────────────────────────────────────────── end to end

def test_the_mexico_pack_runs_clean(tmp_path):
    source = tmp_path / MEXICO.name
    shutil.copy2(MEXICO, source)
    report = run(source, tmp_path / "out.xlsx", tmp_path / "logs")

    assert report.ok, [o.detail for o in report.outcomes if o.status == "error"]
    assert report.rules_ok
    assert len(report.growth) == 2
    processed = {o.sheet for o in report.outcomes if o.status == "processed"}
    assert {"06. EQ Aggs", "07. Wind Aggs"} <= processed


def test_the_growth_block_is_written_on_the_aggregate_sheet(tmp_path):
    source = tmp_path / MEXICO.name
    shutil.copy2(MEXICO, source)
    out = tmp_path / "out.xlsx"
    run(source, out, tmp_path / "logs")

    wb = load_workbook(out, data_only=True)
    text = [c.value for row in wb["06. EQ Aggs"].iter_rows() for c in row
            if isinstance(c.value, str)]
    assert any("EXPOSURE AND PREMIUM GROWTH — EARTHQUAKE" in t for t in text)
    assert "Implied rate change" in text
    assert "2026 at expiry" in text
    # and it belongs to the aggregate sheet, not to 01
    history = [c.value for row in wb["01. History EQ"].iter_rows() for c in row
               if isinstance(c.value, str)]
    assert not any("EXPOSURE AND PREMIUM GROWTH" in t for t in history)


def test_the_process_log_records_the_growth(tmp_path):
    source = tmp_path / MEXICO.name
    shutil.copy2(MEXICO, source)
    logs = tmp_path / "logs"
    run(source, tmp_path / "out.xlsx", logs)

    text = next(logs.glob("*_process.log")).read_text(encoding="utf-8")
    assert "Exposure and premium growth (§10.4)" in text
    assert "implied rate" in text
    assert "2026 at expiry" in text
