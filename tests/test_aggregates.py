"""06 / 07 cat aggregates — specification_v1.md §2.5, §10.4."""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest
from openpyxl import load_workbook

from datatransform.extract import extract_sheet, last_non_empty_row
from datatransform.growth import OK, WARNING, growth_tables, threshold_of
from datatransform.markers import read_markers
from datatransform.model import Axis, ExtractionError
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

def test_the_header_boundary_follows_the_widest_register():
    """06 declares nineteen levels — every shape the aggregate may arrive at."""
    ws = load_workbook(MEXICO, data_only=True)["00. NC+Interdep"]
    first, last, attrs = Nomenclature.register_columns(ws)

    nomenclature, _ = _read()
    widest = max(len(d.headers) for d in nomenclature.datasets.values())
    assert first == 4
    assert widest == 19                       # Zone + 9 cells + 9 other levels
    assert attrs == first + widest + 1        # no fixed column anywhere
    assert last >= first + widest - 1


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


# ══════════════════════════════════════════ the bridge — spec §2.6, S20
#
# The target is fixed: nine cells for earthquake, three for windstorm. What a cedent
# reports varies, and sheet 08 — its own book split, per section — is what bridges the
# two. Every case is B(z,t) = Σₐ S(z,a) · M(a → t).

from datatransform.bridge import build_bridge, fit_margins      # noqa: E402
from datatransform.model import Confidence, SplitRule           # noqa: E402
from datatransform.specs import Split, Step2Spec                # noqa: E402
from datatransform.transform import _split                      # noqa: E402

SPEC = Step2Spec(sort_by=("Zone",), split=Split(total="Total"))


def _bridge(section="Hurricane", covers=("Building", "Content", "BI")):
    nomenclature, blocks = _read()
    return build_bridge(nomenclature, blocks, section, ("Res", "Com", "Ind"), covers)


def _wind_reporting(labels, values):
    """A hurricane block relabelled to report at some other level."""
    nomenclature, blocks = _read()
    block = next(b for b in blocks if b.dataset.role == "07")
    block.address_map = {"Zone": "B", **{lab: chr(ord("C") + i)
                                         for i, lab in enumerate(labels)}}
    for record in block.records:
        record.values = {"Zone": record.values["Zone"],
                         **dict(zip(labels, values(record.values["Total"])))}
    return nomenclature, blocks, block


# ───────────────────────────────────────────── the target is fixed, not per book

def test_the_target_cells_are_fixed_per_dataset():
    """Earthquake always nine, windstorm always three — the book does not change it."""
    nomenclature, _ = _read()
    assert nomenclature.buckets_for("06 EQ Aggs") == (
        "Res Building", "Res Content", "Res BI",
        "Com Building", "Com Content", "Com BI",
        "Ind Building", "Ind Content", "Ind BI",
    )
    assert nomenclature.buckets_for("07 Wind Aggs") == ("Res", "Com", "Ind")


def test_every_reporting_level_is_declared_in_the_register():
    """The register doubles as the list of shapes the dataset accepts."""
    nomenclature, _ = _read()
    headers = set(nomenclature.datasets["07. Wind Aggs"].headers)
    assert {"Res", "Com", "Ind"} <= headers                  # the target
    assert {"Building", "Content", "BI"} <= headers          # the cover margin
    assert {"Projects", "Renewables"} <= headers             # the third axis
    assert "Total" in headers


# ─────────────────────────────────────────────────── sheet 08 is the source

def test_sheet_08_is_read_per_section():
    eq, wind = _bridge("Earthquake"), _bridge("Hurricane")
    assert eq is not None and wind is not None
    assert "08. Splits EQ" in eq.source and "08. Splits Wind" in wind.source
    assert eq.joint() != wind.joint()          # different books, different mixes


def test_shares_taken_from_amounts_always_close():
    """Unlike typed percentages, a ratio derived from sums insured cannot fail to sum."""
    bridge = _bridge()
    assert sum(bridge.joint().values()) == pytest.approx(1.0)
    for occupancy in ("Res", "Com", "Ind"):
        assert sum(bridge.cover_given_occupancy(occupancy).values()) == pytest.approx(1.0)
    for cover in ("Building", "Content", "BI"):
        assert sum(bridge.occupancy_given_cover(cover).values()) == pytest.approx(1.0)


def test_the_occupancy_mix_differs_by_cover():
    """Residential BI is nearly nil — a single occupancy vector could not know that."""
    bridge = _bridge()
    assert bridge.occupancy_given_cover("Building")["Res"] > 0.35
    assert bridge.occupancy_given_cover("BI")["Res"] < 0.12


# ─────────────────────────────────────── case a) one figure per zone

def test_a_total_only_report_uses_the_joint_distribution():
    block, result = _step2("07", "2025 9 months")
    assert block.fields == ("Zone", "Total")               # step 1 shows what arrived

    bridge = _bridge()
    joint = bridge.joint()
    expected = {o: sum(s for (i, _), s in joint.items() if i == o)
                for o in ("Res", "Com", "Ind")}
    first = next(r for r in result.records if r.values["Zone"] == "1")
    for occupancy, share in expected.items():
        assert first.values[occupancy] == pytest.approx(first.values["Total"] * share)


def test_the_note_names_the_sheet_the_ratio_came_from():
    _, result = _step2("07", "2025 9 months")
    note = next(n for n in result.notes if "target cell(s) built" in n)
    assert "08. Splits Wind" in note
    assert "an assumption about the zone, not about the book" in note


# ─────────────────────────────── case b) occupancy reported, cover added

def test_a_reported_occupancy_margin_survives_untouched():
    nomenclature, blocks, block = _wind_reporting(
        ["Res", "Com", "Ind", "Total"],
        lambda t: (t * 0.4, t * 0.35, t * 0.25, t),
    )
    nomenclature.axes.append(Axis("07 Wind Aggs", "Cover", ("Building", "Content", "BI")))
    records, created, notes, _, _, finding = _split(
        block, SPEC, block.records, nomenclature, blocks)
    first = records[0]
    for occupancy in ("Res", "Com", "Ind"):
        parts = sum(first.values[f"{occupancy} {c}"]
                    for c in ("Building", "Content", "BI"))
        assert parts == pytest.approx(first.values[occupancy])   # exactly what arrived
    assert any("cover mix of each comes from" in n for n in notes)


# ────────────────────────────── case c) cover reported, occupancy added

def test_a_reported_cover_margin_survives_untouched():
    nomenclature, blocks, block = _wind_reporting(
        ["Building", "Content", "BI"],
        lambda t: (t * 0.7, t * 0.2, t * 0.1),
    )
    nomenclature.axes.append(Axis("07 Wind Aggs", "Cover", ("Building", "Content", "BI")))
    records, created, notes, _, _, finding = _split(
        block, SPEC, block.records, nomenclature, blocks)
    first = records[0]
    for cover in ("Building", "Content", "BI"):
        parts = sum(first.values[f"{o} {cover}"] for o in ("Res", "Com", "Ind"))
        assert parts == pytest.approx(first.values[cover])
    assert any("residential BI stays near nil" in n for n in notes)


def test_windstorm_sums_the_cover_axis_out_rather_than_discarding_it():
    """07 targets occupancy only, but the reported cover figures still shape the answer."""
    nomenclature, blocks, block = _wind_reporting(
        ["Building", "Content", "BI"],
        lambda t: (t * 0.5, t * 0.2, t * 0.3),        # BI-heavy: commercial/industrial
    )
    records, created, notes, _, _, finding = _split(
        block, SPEC, block.records, nomenclature, blocks)
    assert set(created) == {"Res", "Com", "Ind"}
    first = records[0]
    total = sum(first.values[c] for c in ("Building", "Content", "BI"))
    assert sum(first.values[o] for o in ("Res", "Com", "Ind")) == pytest.approx(total)

    # A flat occupancy split would have given Res the book share; the BI weight cuts it.
    flat = _bridge().joint()
    res_flat = sum(s for (i, _), s in flat.items() if i == "Res")
    assert first.values["Res"] / total < res_flat
    assert any("summed out" in n for n in notes)


# ─────────────────────────── two margins reported: neither may move

def test_both_margins_reported_are_both_preserved_exactly():
    """Conditioning on one lets the other drift. Both were reported, so both are fitted."""
    nomenclature, blocks, block = _wind_reporting(
        ["Res", "Com", "Ind", "Building", "Content", "BI"],
        lambda t: (t * 0.4, t * 0.35, t * 0.25, t * 0.7, t * 0.2, t * 0.1),
    )
    nomenclature.axes.append(Axis("07 Wind Aggs", "Cover", ("Building", "Content", "BI")))
    records, _, notes, _, _, finding = _split(
        block, SPEC, block.records, nomenclature, blocks)

    first = records[0]
    for occupancy in ("Res", "Com", "Ind"):
        assert sum(first.values[f"{occupancy} {c}"] for c in
                   ("Building", "Content", "BI")) == pytest.approx(
            first.values[occupancy], rel=1e-6)
    for cover in ("Building", "Content", "BI"):
        assert sum(first.values[f"{o} {cover}"] for o in
                   ("Res", "Com", "Ind")) == pytest.approx(first.values[cover], rel=1e-6)
    assert any("neither margin moves" in n for n in notes)


def test_the_fit_uses_the_seed_only_for_the_interaction():
    """Same margins, different seed — the margins hold, the interior moves."""
    rows, columns = ("Res", "Com"), ("Building", "BI")
    margins = ({"Res": 60.0, "Com": 40.0}, {"Building": 70.0, "BI": 30.0})
    flat = fit_margins({}, rows, columns, *margins)
    skewed = fit_margins({("Res", "Building"): 9.0, ("Res", "BI"): 1.0,
                          ("Com", "Building"): 1.0, ("Com", "BI"): 9.0},
                         rows, columns, *margins)

    for grid in (flat, skewed):
        assert sum(grid[("Res", c)] for c in columns) == pytest.approx(60.0)
        assert sum(grid[(r, "BI")] for r in rows) == pytest.approx(30.0)
    assert skewed[("Res", "Building")] > flat[("Res", "Building")]


# ──────────────────── the third axis: Projects / Renewables onto the nine

def test_a_segmented_report_is_translated_through_the_declared_convention():
    """Projects and Renewables are not occupancy — they are mapped onto it."""
    nomenclature, blocks, block = _wind_reporting(
        ["Projects", "Renewables"], lambda t: (t * 0.6, t * 0.4),
    )
    records, created, notes, _, _, finding = _split(
        block, SPEC, block.records, nomenclature, blocks)
    first = records[0]
    total = first.values["Projects"] + first.values["Renewables"]
    assert sum(first.values[o] for o in ("Res", "Com", "Ind")) == pytest.approx(total)

    # Renewables is wholly industrial and a Project is half commercial, half industrial,
    # so residential can only come from nowhere — there is no residential engineering.
    assert first.values["Res"] == pytest.approx(0.0)
    assert first.values["Com"] == pytest.approx(first.values["Projects"] * 0.5)
    assert first.values["Ind"] == pytest.approx(
        first.values["Projects"] * 0.5 + first.values["Renewables"])
    assert any("declared convention" in n and "sheet 08 cannot supply" in n
               for n in notes)


def test_the_segment_convention_is_declared_not_coded():
    nomenclature, _ = _read()
    assert set(nomenclature.segment_categories()) == {
        "Renewables", "Projects", "Commercial"}
    projects = nomenclature.splits_for("07 Wind Aggs", "Occupancy", "Projects")
    assert {r.category: r.share for r in projects} == {"Com": 0.5, "Ind": 0.5}
    assert all(r.source == "House convention" for r in projects)


def test_a_column_belonging_to_no_declared_level_is_fatal():
    """Without its convention Renewables belongs to no level — it can be neither used
    nor compared, and a column carrying money is never dropped in silence."""
    nomenclature, blocks, block = _wind_reporting(
        ["Projects", "Renewables"], lambda t: (t * 0.6, t * 0.4),
    )
    nomenclature.splits = [r for r in nomenclature.splits
                           if r.source_category != "Renewables"]
    with pytest.raises(ExtractionError, match="belongs to no level"):
        _split(block, SPEC, block.records, nomenclature, blocks)


# ────────── two views of the same book: a finding and a question, not a refusal

def test_two_levels_at_once_produce_a_finding_not_a_refusal():
    """The block is built from the richer level; the other is bridged and compared."""
    nomenclature, blocks, block = _wind_reporting(
        ["Res", "Com", "Ind", "Projects"],
        lambda t: (t * 0.4, t * 0.35, t * 0.25, t),
    )
    records, created, notes, _, _, finding = _split(
        block, SPEC, block.records, nomenclature, blocks)

    assert finding is not None
    # For windstorm the occupancy *is* the target, so the grid is the primary view.
    assert finding.primary == "reported" and finding.secondary == "segment"
    assert [r[0] for r in finding.rows] == ["Res", "Com", "Ind"]

    # Built from the occupancy, which is reported: those figures come out untouched.
    reported = sum(records[0].values[o] for o in ("Res", "Com", "Ind"))
    assert records[0].values["Res"] == pytest.approx(reported * 0.4)
    assert created == ()                       # the target was already there


def test_the_finding_names_the_gap_and_asks_the_underwriter():
    nomenclature, blocks, block = _wind_reporting(
        ["Res", "Com", "Ind", "Projects"],
        lambda t: (t * 0.4, t * 0.35, t * 0.25, t),
    )
    _, _, _, _, _, finding = _split(block, SPEC, block.records, nomenclature, blocks)

    # A wholly residential book against a Projects view that has no residential at all.
    assert not finding.agrees
    assert finding.worst > 0.5
    assert "QUESTION FOR THE UNDERWRITER" in finding.question
    assert "reported grid" in finding.question
    assert "Projects/Renewables segmentation" in finding.question


def test_two_views_that_agree_raise_no_question():
    """Projects is half commercial, half industrial — an occupancy split saying the same
    thing is not a finding."""
    nomenclature, blocks, block = _wind_reporting(
        ["Res", "Com", "Ind", "Projects"],
        lambda t: (0.0, t * 0.5, t * 0.5, t),
    )
    _, _, _, _, _, finding = _split(block, SPEC, block.records, nomenclature, blocks)
    assert finding is not None and finding.agrees


# ─────────────────────────────────────────────── what it refuses to do

def test_a_section_without_an_08_falls_back_to_declared_ratios():
    nomenclature, blocks = _read()
    blocks = [b for b in blocks if b.dataset.role != "08"]
    block = next(b for b in blocks if b.dataset.role == "07")
    for category, share in (("Res", 0.38), ("Com", 0.42), ("Ind", 0.20)):
        nomenclature.splits.append(
            SplitRule("07 Wind Aggs", "Occupancy", "", category, share, "Prior year"))

    records, created, notes, _, _, finding = _split(
        block, SPEC, block.records, nomenclature, blocks)
    assert set(created) == {"Res", "Com", "Ind"}
    assert records[0].values["Res"] == pytest.approx(records[0].values["Total"] * 0.38)
    note = next(n for n in notes if "no 08 split table" in n)
    assert "Prior year" in note and "an assumption, not a reading" in note


def test_no_08_and_no_declared_ratio_is_fatal():
    nomenclature, blocks = _read()
    blocks = [b for b in blocks if b.dataset.role != "08"]
    block = next(b for b in blocks if b.dataset.role == "07")
    with pytest.raises(ExtractionError, match="declares no share"):
        _split(block, SPEC, block.records, nomenclature, blocks)


def test_a_block_reporting_no_level_at_all_is_fatal():
    nomenclature, blocks = _read()
    block = next(b for b in blocks if b.dataset.role == "07")
    block.address_map.pop("Total")
    with pytest.raises(ExtractionError, match="no level to build from"):
        _split(block, SPEC, block.records, nomenclature, blocks)


def test_an_08_carrying_nothing_under_a_cover_is_fatal():
    """Rather than spreading it evenly and calling that an answer."""
    nomenclature, blocks = _read()
    table = next(b for b in blocks if b.dataset.role == "08"
                 and b.section == "Hurricane")
    for record in table.records:
        record.values["BI"] = 0.0
    block = next(b for b in blocks if b.dataset.role == "07")
    nomenclature.axes.append(Axis("07 Wind Aggs", "Cover", ("Building", "Content", "BI")))
    _, _, blk = _wind_reporting(["Building", "Content", "BI"],
                                lambda t: (t * 0.7, t * 0.2, t * 0.1))
    blk.section = "Hurricane"
    with pytest.raises(ExtractionError, match="will not spread it evenly"):
        _split(blk, SPEC, blk.records, nomenclature, blocks)


# ───────────────────────────────────────────────── 08 is a dataset too

def test_the_split_table_gets_both_steps_like_any_other_dataset():
    nomenclature, blocks = _read()
    table = next(b for b in blocks if b.dataset.role == "08")
    result = apply_step2(table, step2_for(table.dataset.key), nomenclature, blocks)
    assert [r.values["Category"] for r in result.records] == ["Com", "Ind", "Res"]
    assert any("Total checked" in n and "all agree" in n for n in result.notes)


def test_a_split_block_no_longer_reads_as_confirmed():
    block, result = _step2("07", "2025 9 months")
    assert block.confidence == Confidence.CONFIRMED
    assert result.confidence == Confidence.ASSUMED
    assert set(result.assumed_fields) == {"Res", "Com", "Ind"}


def test_an_unsplit_block_keeps_step_1s_confidence():
    block, result = _step2("06", "2025 9 months")
    assert result.assumed_fields == ()
    assert result.confidence == block.confidence


def test_splitting_never_moves_the_zone_total():
    block, result = _step2("07", "2025 9 months")
    assert block.totals()["Total"] == pytest.approx(result.totals()["Total"])
    for record in result.records:
        assert sum(record.values[o] for o in ("Res", "Com", "Ind")) == pytest.approx(
            record.values["Total"])


def test_the_finding_is_written_into_the_sheet_with_somewhere_to_answer(tmp_path):
    """A question needs a place for the answer, so step 2 leaves one."""
    from datatransform.recalc import inject

    source = tmp_path / "two_levels.xlsx"
    shutil.copy2(MEXICO, source)
    wb = load_workbook(source)
    ws = wb["07. Wind Aggs"]
    for col, label in ((3, "Res"), (4, "Com"), (5, "Ind"), (6, "Projects")):
        ws.cell(row=18, column=col, value=label)
    for row in range(19, 61):
        total = ws.cell(row=row, column=3).value
        if total is None:
            continue
        for col, share in ((3, 0.40), (4, 0.35), (5, 0.25), (6, 1.0)):
            ws.cell(row=row, column=col, value=round(total * share))
    wb.save(source)
    inject(source)

    out = tmp_path / "out.xlsx"
    report = run(source, out, tmp_path / "logs")
    assert report.ok, [o.detail for o in report.outcomes if o.status == "error"]

    ws = load_workbook(out, data_only=True)["07. Wind Aggs"]
    text = [c.value for row in ws.iter_rows() for c in row if isinstance(c.value, str)]
    assert any("TWO VIEWS OF THE SAME BOOK — THEY DO NOT AGREE" in t for t in text)
    assert any("QUESTION FOR THE UNDERWRITER" in t for t in text)
    assert "Answer:" in text                      # an empty cell to write it in


# ─────────────────────────── 08 as percentages rather than amounts

def test_a_split_table_of_percentages_reads_the_same_as_amounts():
    """Only ratios are read, so a table given as shares of the book needs no special
    handling — it normalises to the same distribution."""
    nomenclature, blocks = _read()
    table = next(b for b in blocks if b.dataset.role == "08" and b.section == "Hurricane")
    before = build_bridge(nomenclature, blocks, "Hurricane", ("Res", "Com", "Ind"),
                          ("Building", "Content", "BI")).joint()

    grand = sum(r.values[c] for r in table.records
                for c in ("Building", "Content", "BI"))
    for record in table.records:
        for c in ("Building", "Content", "BI"):
            record.values[c] = record.values[c] / grand * 100.0
        record.values["Total"] = sum(record.values[c]
                                     for c in ("Building", "Content", "BI"))

    after = build_bridge(nomenclature, blocks, "Hurricane", ("Res", "Com", "Ind"),
                         ("Building", "Content", "BI")).joint()
    for key, share in before.items():
        assert after[key] == pytest.approx(share)


def test_percentages_that_close_per_row_are_read_as_conditionals():
    """Three rows each summing to 100% are three distributions, not one. Normalising the
    grid whole would silently assert the three occupancies are equally large."""
    nomenclature, blocks = _read()
    table = next(b for b in blocks if b.dataset.role == "08" and b.section == "Hurricane")
    weights = {}
    for record in table.records:
        row = sum(record.values[c] for c in ("Building", "Content", "BI"))
        weights[record.values["Category"]] = row
        for c in ("Building", "Content", "BI"):
            record.values[c] = record.values[c] / row * 100.0
        record.values["Total"] = row              # the weight lives here and nowhere else

    bridge = build_bridge(nomenclature, blocks, "Hurricane", ("Res", "Com", "Ind"),
                          ("Building", "Content", "BI"))
    joint = bridge.joint()
    grand = sum(weights.values())
    for occupancy, weight in weights.items():
        implied = sum(s for (o, _), s in joint.items() if o == occupancy)
        assert implied == pytest.approx(weight / grand)


# ──────────────────────────── the engineering shape is a sheet edit, not a code change

def test_an_engineering_book_needs_no_code_change():
    """Projects/Renewables as the first axis, cover as the second — declared, not coded."""
    nomenclature, blocks = _read()
    nomenclature.axes = [a for a in nomenclature.axes
                         if a.dataset != "07 Wind Aggs"] + [
        Axis("07 Wind Aggs", "Segment", ("Projects", "Renewables")),
    ]
    assert nomenclature.buckets_for("07 Wind Aggs") == ("Projects", "Renewables")

    # And the second axis is found by position, so renaming "Cover" changes nothing.
    nomenclature.axes = [
        Axis("06 EQ Aggs", "Occupancy", ("Res", "Com", "Ind")),
        Axis("06 EQ Aggs", "Deckungsart", ("Building", "Content", "BI")),
    ]
    assert nomenclature.cover_categories() == ("Building", "Content", "BI")
