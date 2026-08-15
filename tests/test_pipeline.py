"""Tests for the rules in specification_v1.md."""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest
from openpyxl import load_workbook

from datatransform.extract import extract_sheet, last_non_empty_row
from datatransform.markers import attributes_for, block_indices, parse_marker, read_markers
from datatransform.model import Confidence, ExtractionError, Orientation
from datatransform.nomenclature import match_sheet, norm, read_nomenclature, sheet_sort_key
from datatransform.runner import run
from datatransform.specs import step2_for
from datatransform.transform import apply_step2
from datatransform.writer import ANCHOR_PREFIX

SOURCE = Path(__file__).resolve().parents[1] / "Intake_v1.xlsx"
ROW_WISE = "01. History"
TRANSPOSED = "01. History_Transposed"

EXPECTED = [
    (2021, 15900, 14930),
    (2022, 16740, 10030),
    (2023, 17520, 12660),
    (2024, 18390, 11700),
    (2025, 19200, 10250),
]


@pytest.fixture(scope="module")
def books():
    return (load_workbook(SOURCE, data_only=True), load_workbook(SOURCE, data_only=False))


@pytest.fixture(scope="module")
def nomenclature(books):
    return read_nomenclature(books[1])


def _block(books, nomenclature, sheet):
    values, formulas = books
    markers = read_markers(formulas[sheet], last_non_empty_row(values[sheet]))
    blocks = extract_sheet(values[sheet], formulas[sheet], nomenclature, markers)
    assert len(blocks) == 1
    return blocks[0]


# ───────────────────────────────────────────────────────── §4 marker grammar

@pytest.mark.parametrize("text,name,value,index,hyp", [
    ("Header_1", "Header", None, 1, False),
    ("Info_1 = L", "Info", "L", 1, False),
    ("Info_1 = 16", "Info", "16", 1, False),
    ("Transpose_1", "Transpose", None, 1, False),
    ("Currency = USD", "Currency", "USD", None, False),
    ("H_Year basis = UW", "Year basis", "UW", None, True),
    ("H_PF transfer = with clean cut", "PF transfer", "with clean cut", None, True),
    ("Currency_2 = EUR", "Currency", "EUR", 2, False),
])
def test_marker_grammar(text, name, value, index, hyp):
    m = parse_marker(text, 1)
    assert (m.name, m.value, m.index, m.is_hypothesis) == (name, value, index, hyp)


def test_block_declared_only_by_header(books, nomenclature):
    values, formulas = books
    markers = read_markers(formulas[ROW_WISE])
    assert block_indices(markers) == [1]


def test_block_override_beats_sheet_wide_attribute():
    markers = [parse_marker("Currency = USD", 1), parse_marker("Currency_2 = EUR", 2)]
    assert attributes_for(markers, 1)["Currency"].value == "USD"
    assert attributes_for(markers, 2)["Currency"].value == "EUR"


def test_no_header_marker_extracts_nothing(books, nomenclature):
    values, formulas = books
    blocks = extract_sheet(values[ROW_WISE], formulas[ROW_WISE], nomenclature, markers=[])
    assert blocks == []


# ───────────────────────────────────────────────────── §2.2, §8 name matching

def test_numeric_prefix_matching_survives_misspelling():
    assert sheet_sort_key("10. Triangles")[0] == sheet_sort_key("10. Triangels")[0]
    assert match_sheet("10. Triangles", ["10. Triangels"]) == "10. Triangels"


def test_norm_strips_nbsp_and_quote_style():
    assert norm("  Premium  ") == "Premium"
    assert norm("„01. History“") == '"01. History"'


# ───────────────────────────────────────────────────────── §3 nomenclature

def test_nomenclature_declares_dataset_01(nomenclature):
    ds = nomenclature.datasets[ROW_WISE]
    assert ds.key == "01 History"
    assert ds.headers == ("Year", "Premium", "Incurred Losses")
    assert ds.key_field == "Year"
    assert ds.measures == ("Premium", "Incurred Losses")
    assert "PF transfer" in ds.attributes


def test_vocabulary_is_enforced_only_where_defined(nomenclature):
    assert nomenclature.validate_value("Year basis", "UW")
    assert not nomenclature.validate_value("Year basis", "Underwriting")
    assert nomenclature.validate_value("Currency", "USD")
    assert not nomenclature.validate_value("Currency", "Dollars")
    assert nomenclature.validate_value("As at", "31.12.2025")     # no vocabulary entry


# ───────────────────────────────────────────────── §7, §8 resolution & selection

def test_row_wise_resolution(books, nomenclature):
    b = _block(books, nomenclature, ROW_WISE)
    assert b.orientation is Orientation.ROW_WISE
    assert b.header_ref == "8" and b.info_ref == "L"
    assert b.address_map == {"Year": "B", "Premium": "D", "Incurred Losses": "G"}
    assert b.provenance_label == "Source row"


def test_transposed_resolution(books, nomenclature):
    b = _block(books, nomenclature, TRANSPOSED)
    assert b.orientation is Orientation.TRANSPOSED
    assert b.header_ref == "C" and b.info_ref == "16"
    assert b.address_map == {"Year": "8", "Premium": "10", "Incurred Losses": "13"}
    assert b.provenance_label == "Source column"


@pytest.mark.parametrize("sheet", [ROW_WISE, TRANSPOSED])
def test_records_and_exclusion(books, nomenclature, sheet):
    b = _block(books, nomenclature, sheet)
    assert (b.candidates, len(b.records), b.excluded) == (6, 5, 1)
    got = [(r.values["Year"], r.values["Premium"], r.values["Incurred Losses"])
           for r in b.records]
    assert got == EXPECTED


def test_both_orientations_are_equivalent(books, nomenclature):
    a = _block(books, nomenclature, ROW_WISE)
    b = _block(books, nomenclature, TRANSPOSED)
    assert [r.values for r in a.records] == [r.values for r in b.records]
    assert a.totals() == b.totals() == {"Premium": 87750, "Incurred Losses": 59570}
    assert a.provenance_label != b.provenance_label


def test_excluded_total_ties_to_extracted_total(books, nomenclature):
    b = _block(books, nomenclature, ROW_WISE)
    assert len(b.excluded_records) == 1
    excluded = b.excluded_records[0].values
    assert excluded["Premium"] == b.totals()["Premium"]
    assert excluded["Incurred Losses"] == b.totals()["Incurred Losses"]


def test_undeclared_columns_are_logged_not_dropped(books, nomenclature):
    assert _block(books, nomenclature, ROW_WISE).unextracted == ["C", "E", "F"]
    assert _block(books, nomenclature, TRANSPOSED).unextracted == ["9", "11", "12"]


def test_only_declared_addresses_are_read(books, nomenclature):
    b = _block(books, nomenclature, ROW_WISE)
    assert len(b.records) * len(b.dataset.headers) == 15
    assert all(set(r.values) == set(b.dataset.headers) for r in b.records)


# ───────────────────────────────────────────────────────── §5 confidence

def test_confidence_is_derived_and_worst_wins(books, nomenclature):
    b = _block(books, nomenclature, ROW_WISE)
    assert {h.attribute for h in b.hypotheses} == {"Year basis", "Loss basis", "PF transfer"}
    assert all(h.confidence is Confidence.ASSUMED for h in b.hypotheses)
    assert b.confidence is Confidence.ASSUMED


def test_open_ranks_worse_than_assumed():
    assert Confidence.OPEN.rank > Confidence.ASSUMED.rank > Confidence.CONFIRMED.rank
    assert Confidence.worst([Confidence.CONFIRMED, Confidence.OPEN]) is Confidence.OPEN
    assert Confidence.worst([]) is Confidence.CONFIRMED


def test_missing_attribute_raises_open_hypothesis(books, nomenclature, monkeypatch):
    values, formulas = books
    ds = nomenclature.datasets[ROW_WISE]
    markers = [m for m in read_markers(formulas[ROW_WISE], last_non_empty_row(values[ROW_WISE]))
               if m.name != "PF transfer"]
    blocks = extract_sheet(values[ROW_WISE], formulas[ROW_WISE], nomenclature, markers)
    opens = [h for h in blocks[0].hypotheses if h.confidence is Confidence.OPEN]
    assert [h.attribute for h in opens] == ["PF transfer"]
    assert blocks[0].confidence is Confidence.OPEN


def test_row_level_confidence_only_when_transposed(books, nomenclature):
    assert all(r.confidence is None for r in _block(books, nomenclature, ROW_WISE).records)
    assert all(r.confidence is Confidence.ASSUMED
               for r in _block(books, nomenclature, TRANSPOSED).records)


# ───────────────────────────────────────────────────────── §9.2 step 2

def test_step2_sorts_reorders_and_calculates(books, nomenclature):
    b = _block(books, nomenclature, ROW_WISE)
    result = apply_step2(b, step2_for(b.dataset.key))
    assert [r.values["Year"] for r in result.records] == [2021, 2022, 2023, 2024, 2025]
    assert result.columns == ("Year", "Premium", "Incurred Losses", "Loss Ratio %")
    ratios = result.computed["Loss Ratio %"]
    assert ratios[0] == pytest.approx(14930 / 15900)
    assert ratios[-1] == pytest.approx(10250 / 19200)


def test_step2_is_value_preserving(books, nomenclature):
    b = _block(books, nomenclature, ROW_WISE)
    result = apply_step2(b, step2_for(b.dataset.key))
    assert result.totals() == b.totals()


def test_step2_guard_catches_a_value_changing_bug(books, nomenclature, monkeypatch):
    """Sorting cannot move a total; if it does, that is a bug — spec §10 C3."""
    from datatransform import transform as tmod

    b = _block(books, nomenclature, ROW_WISE)
    monkeypatch.setattr(tmod, "sorted", lambda seq, **kw: list(seq)[:-1], raising=False)
    with pytest.raises(ExtractionError, match="value-preserving"):
        apply_step2(b, step2_for(b.dataset.key))


# ───────────────────────────────────────────────────────── §9 end to end

@pytest.fixture
def workspace(tmp_path):
    src = tmp_path / "Intake.xlsx"
    shutil.copy2(SOURCE, src)
    return src


def test_run_writes_blocks_and_leaves_source_untouched(workspace, tmp_path):
    before = workspace.read_bytes()
    out = tmp_path / "out.xlsx"
    report = run(workspace, out, tmp_path / "logs")

    assert report.ok
    assert workspace.read_bytes() == before, "the source workbook must never be modified"
    assert out.exists()

    wb = load_workbook(out, data_only=True)
    anchors = [
        c.value for row in wb[ROW_WISE].iter_rows() for c in row
        if isinstance(c.value, str) and c.value.startswith(ANCHOR_PREFIX)
    ]
    assert anchors == ["⟦DT:01 History:STEP1:v1⟧", "⟦DT:01 History:STEP2:v1⟧"]


def test_generated_blocks_start_below_the_last_original_row(workspace, tmp_path):
    original_last = last_non_empty_row(load_workbook(workspace, data_only=True)[ROW_WISE])
    out = tmp_path / "out.xlsx"
    run(workspace, out, tmp_path / "logs")

    ws = load_workbook(out, data_only=True)[ROW_WISE]
    anchor = next(c.row for row in ws.iter_rows() for c in row
                  if isinstance(c.value, str) and c.value.startswith(ANCHOR_PREFIX))
    assert anchor == original_last + 1 + 3          # three blank rows — spec §9.1 O2


def test_rerun_replaces_rather_than_stacks(workspace, tmp_path):
    """Processing a workbook that already carries generated blocks must replace them."""
    out = tmp_path / "out.xlsx"
    run(workspace, out, tmp_path / "logs")
    first_extent = load_workbook(out, data_only=True)[ROW_WISE].max_row

    already_processed = tmp_path / "again_in.xlsx"
    shutil.copy2(out, already_processed)
    again = tmp_path / "again_out.xlsx"
    run(already_processed, again, tmp_path / "logs")

    ws = load_workbook(again, data_only=True)[ROW_WISE]
    anchors = [c.value for row in ws.iter_rows() for c in row
               if isinstance(c.value, str) and c.value.startswith(ANCHOR_PREFIX)]
    assert len(anchors) == 2, "a re-run must replace the previous blocks, not append to them"
    assert ws.max_row == first_extent


def test_output_may_not_overwrite_the_source(workspace, tmp_path):
    with pytest.raises(ExtractionError, match="never modified"):
        run(workspace, workspace, tmp_path / "logs")


def test_formulas_carry_cached_values(workspace, tmp_path):
    out = tmp_path / "out.xlsx"
    run(workspace, out, tmp_path / "logs")
    values = load_workbook(out, data_only=True)[ROW_WISE]
    formulas = load_workbook(out, data_only=False)[ROW_WISE]

    checked = 0
    for row in formulas.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cached = values[cell.coordinate].value
                assert cached is not None, f"{cell.coordinate} {cell.value} has no cached value"
                checked += 1
    assert checked >= 20


def test_control_checks_all_read_ok(workspace, tmp_path):
    out = tmp_path / "out.xlsx"
    run(workspace, out, tmp_path / "logs")
    wb = load_workbook(out, data_only=True)
    for sheet in (ROW_WISE, TRANSPOSED):
        ws = wb[sheet]
        checks = [c.value for row in ws.iter_rows() for c in row if c.value == "MISMATCH"]
        assert not checks
        oks = [c.value for row in ws.iter_rows() for c in row if c.value == "OK"]
        assert len(oks) == 4                        # two measures × step 1 and step 2


def test_logs_are_written(workspace, tmp_path):
    logs = tmp_path / "logs"
    run(workspace, tmp_path / "out.xlsx", logs)
    assert len(list(logs.glob("*_debug.log"))) == 1
    process = next(logs.glob("*_process.log")).read_text(encoding="utf-8")
    assert "sha256" in process
    assert "Fields resolved by label" in process
    assert "Dataset confidence: Assumed" in process


# ───────────────────────────────────────────────── §7.2, §8 fatal conditions

def test_uncalculated_selector_is_fatal(workspace, tmp_path, nomenclature):
    wb = load_workbook(workspace, data_only=False)
    values = load_workbook(workspace, data_only=True)
    values[ROW_WISE]["L10"] = None                  # cache gone, formula still present
    markers = read_markers(wb[ROW_WISE], last_non_empty_row(values[ROW_WISE]))
    with pytest.raises(ExtractionError, match="workbook not calculated"):
        extract_sheet(values[ROW_WISE], wb[ROW_WISE], nomenclature, markers)


def test_broken_selector_integrity_is_fatal(workspace, nomenclature):
    wb = load_workbook(workspace, data_only=False)
    values = load_workbook(workspace, data_only=True)
    values[ROW_WISE]["L10"] = 99                    # says row 99, sits on row 10
    markers = read_markers(wb[ROW_WISE], last_non_empty_row(values[ROW_WISE]))
    with pytest.raises(ExtractionError, match="manipulated"):
        extract_sheet(values[ROW_WISE], wb[ROW_WISE], nomenclature, markers)


def test_error_cell_is_fatal_not_coerced(workspace, nomenclature):
    wb = load_workbook(workspace, data_only=False)
    values = load_workbook(workspace, data_only=True)
    values[ROW_WISE]["D10"] = "#REF!"
    markers = read_markers(wb[ROW_WISE], last_non_empty_row(values[ROW_WISE]))
    with pytest.raises(ExtractionError, match="#REF!"):
        extract_sheet(values[ROW_WISE], wb[ROW_WISE], nomenclature, markers)


def test_missing_declared_label_is_fatal(workspace, nomenclature):
    wb = load_workbook(workspace, data_only=False)
    values = load_workbook(workspace, data_only=True)
    values[ROW_WISE]["D8"] = None                   # remove 'Premium' from extraction row
    markers = read_markers(wb[ROW_WISE], last_non_empty_row(values[ROW_WISE]))
    with pytest.raises(ExtractionError, match="absent from the extraction"):
        extract_sheet(values[ROW_WISE], wb[ROW_WISE], nomenclature, markers)


def test_duplicate_label_is_fatal(workspace, nomenclature):
    wb = load_workbook(workspace, data_only=False)
    values = load_workbook(workspace, data_only=True)
    values[ROW_WISE]["E8"] = "Premium"              # a second 'Premium'
    markers = read_markers(wb[ROW_WISE], last_non_empty_row(values[ROW_WISE]))
    with pytest.raises(ExtractionError, match="ambiguous"):
        extract_sheet(values[ROW_WISE], wb[ROW_WISE], nomenclature, markers)


def test_value_outside_vocabulary_is_fatal(workspace, nomenclature):
    wb = load_workbook(workspace, data_only=False)
    values = load_workbook(workspace, data_only=True)
    wb[ROW_WISE]["A4"] = "H_Year basis = Underwriting"   # markers are read from the formula handle
    markers = read_markers(wb[ROW_WISE], last_non_empty_row(values[ROW_WISE]))
    with pytest.raises(ExtractionError, match="not in its vocabulary"):
        extract_sheet(values[ROW_WISE], wb[ROW_WISE], nomenclature, markers)
