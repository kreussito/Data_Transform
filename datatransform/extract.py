"""Field resolution and record selection. Specification_v1.md §7, §8."""

from __future__ import annotations

import re

from openpyxl.utils import column_index_from_string as col_idx
from openpyxl.utils import get_column_letter as col_letter

from .coerce import coerce
from .markers import attributes_for, block_indices, structural
from .model import (
    Attribute,
    Block,
    Confidence,
    Dataset,
    ExtractionError,
    FieldType,
    Hypothesis,
    Orientation,
    Record,
)
from .nomenclature import Nomenclature, norm

ERROR_CELLS = {"#REF!", "#N/A", "#VALUE!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!"}
LEADING_NUMBER = re.compile(r"^\s*(\d+)")


def last_non_empty_row(ws) -> int:
    """``max_row`` counts formatted-but-empty cells, so find the real extent — spec §9.1 O2."""
    for row in range(ws.max_row, 0, -1):
        for col in range(1, ws.max_column + 1):
            if ws.cell(row=row, column=col).value not in (None, ""):
                return row
    return 0


def _is_error(value) -> bool:
    return isinstance(value, str) and value.strip() in ERROR_CELLS


def _check_cell(values_ws, formulas_ws, row: int, col: int, where: str):
    """Spec §7.2 (uncalculated formulas) and §12 T2 (error cells)."""
    v = values_ws.cell(row=row, column=col).value
    if _is_error(v):
        raise ExtractionError(
            f"{where}: error cell {col_letter(col)}{row} contains {v!r} — "
            "error values are fatal and are never coerced"
        )
    if v is None:
        f = formulas_ws.cell(row=row, column=col).value
        if isinstance(f, str) and f.startswith("="):
            raise ExtractionError(
                f"{where}: {col_letter(col)}{row} holds formula {f!r} with no cached value — "
                "workbook not calculated. Open and save it in Excel, then re-run."
            )
    return v


def _lenient(value, block: Block, label: str, ref: str):
    """Coerce an *excluded* record, tolerating failure.

    Excluded records are shown only as an independent control, so a value that cannot
    be typed must not stop a run over records that were never extracted.
    """
    try:
        return coerce(value, block.field_types[label], label, ref, [])
    except ExtractionError:
        return value


def _resolve_attributes(block: Block, markers, nomenclature: Nomenclature) -> None:
    """Attributes, hypotheses and derived confidence — spec §5.

    Three tiers: ⟦GLOBAL⟧ in sheet 00, then sheet-wide markers, then block markers.
    """
    found = attributes_for(markers, block.index)

    for name, value in nomenclature.globals.items():
        if name in block.dataset.attributes and name not in found:
            block.attributes[name] = Attribute(name, value, False, 0)

    for name, marker in found.items():
        if name not in block.dataset.attributes:
            block.hypotheses.append(
                Hypothesis(
                    id="",
                    dataset_key=block.dataset.key,
                    attribute=name,
                    value=marker.value,
                    confidence=Confidence.ASSUMED,
                    source="tool",
                    note="attribute declared in the sheet but not listed in sheet 00",
                )
            )
            continue

        if not nomenclature.validate_value(name, marker.value):
            allowed = " | ".join(nomenclature.vocabulary.get(name, []))
            raise ExtractionError(
                f"{block.sheet_name}: attribute {name!r} has value {marker.value!r}, "
                f"which is not in its vocabulary ({allowed})"
            )

        block.attributes[name] = Attribute(name, marker.value, marker.is_hypothesis, marker.row)

    for name in block.dataset.attributes:
        attr = block.attributes.get(name)
        if attr is None:
            block.hypotheses.append(
                Hypothesis(
                    id="",
                    dataset_key=block.dataset.key,
                    attribute=name,
                    value=None,
                    confidence=Confidence.OPEN,
                    source="tool",
                    note="listed in sheet 00 but never declared in the sheet — nobody decided",
                )
            )
        elif attr.is_hypothesis:
            block.hypotheses.append(
                Hypothesis(
                    id="",
                    dataset_key=block.dataset.key,
                    attribute=name,
                    value=attr.value,
                    confidence=Confidence.ASSUMED,
                    source="H_ marker",
                    note=f"declared as a hypothesis at A{attr.row}",
                )
            )


def _check_period_overlap(block: Block) -> None:
    """Repeated leading year means overlapping periods — spec §9.2 S10.

    ``2026`` and ``2026 9 months`` are two records of the same year, so a column total
    counts that year twice. The tool flags it and changes nothing: which record to use
    is an underwriting judgment, not the tool's to make.
    """
    key = block.dataset.key_field
    if block.field_types.get(key) is not FieldType.TEXT:
        return

    groups: dict[str, list[str]] = {}
    for record in block.records:
        label = record.values.get(key)
        if not isinstance(label, str):
            continue
        match = LEADING_NUMBER.match(label)
        if match:
            groups.setdefault(match.group(1), []).append(label)

    for year, labels in sorted(groups.items()):
        if len(labels) > 1:
            block.hypotheses.append(
                Hypothesis(
                    id="",
                    dataset_key=block.dataset.key,
                    attribute=f"Period overlap {year}",
                    value=" · ".join(labels),
                    confidence=Confidence.ASSUMED,
                    source="tool",
                    note=(
                        f"{key} {year} appears {len(labels)} times as overlapping "
                        "periods; the control sum counts each of them, so it verifies "
                        "extraction rather than being a portfolio total"
                    ),
                )
            )


def _number_hypotheses(block: Block) -> None:
    for n, hypothesis in enumerate(block.hypotheses, start=1):
        hypothesis.id = f"H-{n:02d}"


def _resolve_labels(cells: list[tuple], dataset: Dataset, where: str) -> dict[str, str]:
    """Strict label match, duplicates and absences fatal — spec §8 F1–F4.

    Takes a list of ``(text, ref)`` pairs rather than a mapping: keying by cell value
    would collapse duplicates, which are precisely what F3 must detect.
    """
    wanted = {norm(h).casefold(): h for h in dataset.headers}
    hits: dict[str, list[int]] = {}
    for text, ref in cells:
        key = norm(text).casefold()
        if key in wanted:
            hits.setdefault(wanted[key], []).append(ref)

    for label, refs in hits.items():
        if len(refs) > 1:
            raise ExtractionError(
                f"{where}: label {label!r} appears {len(refs)} times in the extraction "
                f"row/column — ambiguous"
            )
    missing = [h for h in dataset.headers if h not in hits]
    if missing:
        raise ExtractionError(
            f"{where}: label(s) {missing} declared in sheet 00 but absent from the "
            f"extraction row/column"
        )
    return {label: str(refs[0]) for label, refs in hits.items()}


def extract_block(values_ws, formulas_ws, dataset: Dataset, markers, index: int,
                  nomenclature: Nomenclature) -> Block:
    header = structural(markers, "Header", index)
    info = structural(markers, "Info", index)
    transposed = structural(markers, "Transpose", index) is not None
    where = f"{values_ws.title} block {index}"

    if info is None or not info.value:
        raise ExtractionError(f"{where}: Header_{index} present but Info_{index} missing")

    orientation = Orientation.TRANSPOSED if transposed else Orientation.ROW_WISE
    limit_row = last_non_empty_row(values_ws)
    limit_col = values_ws.max_column

    if orientation is Orientation.ROW_WISE:
        if header.value:
            raise ExtractionError(
                f"{where}: Header_{index} carries a value but Transpose_{index} is absent"
            )
        header_ref, info_ref = str(header.row), info.value.strip().upper()
        cells = [
            (values_ws.cell(row=header.row, column=c).value, c)
            for c in range(1, limit_col + 1)
            if values_ws.cell(row=header.row, column=c).value is not None
        ]
        address = _resolve_labels(cells, dataset, where)
        address = {label: col_letter(int(ref)) for label, ref in address.items()}
    else:
        if not header.value:
            raise ExtractionError(
                f"{where}: Transpose_{index} present, so Header_{index} must name a column"
            )
        header_ref, info_ref = header.value.strip().upper(), info.value.strip()
        hcol = col_idx(header_ref)
        cells = [
            (values_ws.cell(row=r, column=hcol).value, r)
            for r in range(1, limit_row + 1)
            if values_ws.cell(row=r, column=hcol).value is not None
        ]
        address = _resolve_labels(cells, dataset, where)

    block = Block(
        dataset=dataset,
        sheet_name=values_ws.title,
        index=index,
        orientation=orientation,
        header_ref=header_ref,
        info_ref=info_ref,
        address_map=address,
        field_types=nomenclature.field_types(dataset.headers),
    )
    _resolve_attributes(block, markers, nomenclature)

    if orientation is Orientation.ROW_WISE:
        _select_rows(block, values_ws, formulas_ws, limit_row, limit_col, where)
    else:
        _select_columns(block, values_ws, formulas_ws, limit_row, limit_col, where)

    _check_period_overlap(block)
    _number_hypotheses(block)
    return block


def _select_rows(block: Block, values_ws, formulas_ws, limit_row, limit_col, where):
    sel = col_idx(block.info_ref)
    header_row = int(block.header_ref)
    # A candidate must carry at least one measure; text in the key column (a footnote,
    # a section caption) is not a record that was excluded.
    measure_cols = [col_idx(block.address_map[m]) for m in block.numeric_fields
                    if m in block.address_map]

    # Records follow the extraction row; it is never itself a record — spec §6.3, §7.
    for row in range(header_row + 1, limit_row + 1):
        marker = values_ws.cell(row=row, column=sel).value
        if marker is None:
            formula = formulas_ws.cell(row=row, column=sel).value
            if isinstance(formula, str) and formula.startswith("="):
                raise ExtractionError(
                    f"{where}: selector {block.info_ref}{row} holds {formula!r} with no cached "
                    "value — workbook not calculated. A skipped record would be silent."
                )
            if any(values_ws.cell(row=row, column=c).value is not None
                   for c in measure_cols):
                block.candidates += 1
                block.excluded += 1
                block.excluded_records.append(
                    Record(
                        source_ref=str(row),
                        values={
                            label: _lenient(values_ws.cell(row=row, column=col_idx(col)).value,
                                            block, label, f"{col}{row}")
                            for label, col in block.address_map.items()
                        },
                    )
                )
            continue

        block.candidates += 1
        if marker != row:
            raise ExtractionError(
                f"{where}: selector {block.info_ref}{row} reads {marker!r} but sits on row "
                f"{row} — the sheet has been manipulated and the markers no longer align"
            )
        values = {}
        for label, col in block.address_map.items():
            raw = _check_cell(values_ws, formulas_ws, row, col_idx(col), where)
            values[label] = coerce(raw, block.field_types[label], label,
                                   f"{col}{row}", block.coercions)
        block.records.append(Record(source_ref=str(row), values=values))

    extracted = {col_idx(c) for c in block.address_map.values()}
    rows = [int(r.source_ref) for r in block.records]
    for col in range(2, limit_col + 1):
        if col == sel or col in extracted:
            continue
        if any(values_ws.cell(row=r, column=col).value is not None for r in rows):
            block.unextracted.append(col_letter(col))


def _select_columns(block: Block, values_ws, formulas_ws, limit_row, limit_col, where):
    sel = int(block.info_ref)
    header_col = col_idx(block.header_ref)
    measure_rows = [int(block.address_map[m]) for m in block.numeric_fields
                    if m in block.address_map]

    # Records follow the extraction column; column A is the marker channel — spec §4 M1, §6.3.
    for col in range(header_col + 1, limit_col + 1):
        marker = values_ws.cell(row=sel, column=col).value
        if marker is None:
            formula = formulas_ws.cell(row=sel, column=col).value
            if isinstance(formula, str) and formula.startswith("="):
                raise ExtractionError(
                    f"{where}: selector {col_letter(col)}{sel} holds {formula!r} with no cached "
                    "value — workbook not calculated. A skipped record would be silent."
                )
            if any(values_ws.cell(row=r, column=col).value is not None
                   for r in measure_rows):
                block.candidates += 1
                block.excluded += 1
                block.excluded_records.append(
                    Record(
                        source_ref=col_letter(col),
                        values={
                            label: _lenient(values_ws.cell(row=int(row), column=col).value,
                                            block, label, f"{col_letter(col)}{row}")
                            for label, row in block.address_map.items()
                        },
                    )
                )
            continue

        block.candidates += 1
        if marker != col:
            raise ExtractionError(
                f"{where}: selector {col_letter(col)}{sel} reads {marker!r} but sits in column "
                f"{col_letter(col)} ({col}) — the sheet has been manipulated"
            )
        values = {}
        for label, row in block.address_map.items():
            raw = _check_cell(values_ws, formulas_ws, int(row), col, where)
            values[label] = coerce(raw, block.field_types[label], label,
                                   f"{col_letter(col)}{row}", block.coercions)
        # Row-level confidence applies to transposed blocks only — spec §5.3, user ad 7.
        block.records.append(
            Record(source_ref=col_letter(col), values=values, confidence=block.confidence)
        )

    extracted = {int(r) for r in block.address_map.values()}
    cols = [col_idx(r.source_ref) for r in block.records]
    for row in range(1, limit_row + 1):
        if row == sel or row in extracted:
            continue
        if any(values_ws.cell(row=row, column=c).value is not None for c in cols):
            block.unextracted.append(str(row))


def extract_sheet(values_ws, formulas_ws, nomenclature: Nomenclature, markers) -> list[Block]:
    """All blocks of one sheet. No ``Header_i`` → nothing extracted — spec §4 M7."""
    dataset = nomenclature.dataset_for(values_ws.title)
    if dataset is None:
        return []
    return [
        extract_block(values_ws, formulas_ws, dataset, markers, i, nomenclature)
        for i in block_indices(markers)
    ]
