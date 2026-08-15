"""Writing step blocks beneath the original data. Specification_v1.md §9, §10."""

from __future__ import annotations

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as col_letter

from .model import Block, Confidence
from .specs import SPEC_VERSION
from .transform import Step2Result

FONT = "Arial"
ANCHOR_PREFIX = "⟦DT:"
GAP = 3                       # blank rows between blocks — spec §9.1 O2/O3
FIRST_COL = 2                 # column B; column A stays the marker channel

anchor_f = Font(name=FONT, size=9, bold=True, color="7F6000")
title_f = Font(name=FONT, size=11, bold=True)
note_f = Font(name=FONT, size=9, italic=True, color="595959")
head_f = Font(name=FONT, size=10, bold=True)
body_f = Font(name=FONT, size=10)
ctrl_f = Font(name=FONT, size=10, bold=True)

step1_fill = PatternFill("solid", fgColor="DDEBF7")
step2_fill = PatternFill("solid", fgColor="E2EFDA")
ctrl_fill = PatternFill("solid", fgColor="FFF2CC")


def anchor_tag(key: str, step: str) -> str:
    return f"{ANCHOR_PREFIX}{key}:{step}:v{SPEC_VERSION}⟧"


def clear_generated(ws) -> int:
    """Re-runs replace rather than stack — spec §9.1 O5."""
    anchors = [
        r for r in range(1, ws.max_row + 1)
        if isinstance(ws.cell(row=r, column=FIRST_COL).value, str)
        and ws.cell(row=r, column=FIRST_COL).value.startswith(ANCHOR_PREFIX)
    ]
    if not anchors:
        return 0
    start = max(1, min(anchors) - GAP)
    removed = ws.max_row - start + 1
    if removed > 0:
        ws.delete_rows(start, removed)
    return removed


def _number_format(block: Block, label: str) -> str:
    if label == block.dataset.key_field:
        return "0"
    return "#,##0"


class BlockWriter:
    """Writes the two step blocks and reports the cached values its formulas need."""

    def __init__(self, ws, start_row: int):
        self.ws = ws
        self.row = start_row
        self.formula_values: list[tuple[str, str, float]] = []

    def _put(self, col: int, value, font=body_f, fmt=None, fill=None):
        cell = self.ws.cell(row=self.row, column=col, value=value)
        cell.font = font
        if fmt:
            cell.number_format = fmt
        if fill:
            cell.fill = fill
        return cell

    def _line(self, text, font=note_f, col=FIRST_COL):
        self._put(col, text, font)
        self.row += 1

    def _formula(self, col: int, formula: str, expected, fmt=None, font=ctrl_f, fill=None):
        cell = self._put(col, formula, font, fmt, fill)
        if expected is not None:
            self.formula_values.append((self.ws.title, cell.coordinate, expected))
        return cell

    # ────────────────────────────────────────────────────────── step 1
    def write_step1(self, block: Block) -> None:
        ds = block.dataset
        self._put(FIRST_COL, anchor_tag(ds.key, "STEP1"), anchor_f)
        self.row += 1
        self._put(FIRST_COL, "STEP 1 — FILTERED & INTERPRETED", title_f, fill=step1_fill)
        self.row += 1

        self._line(
            f"Source: {block.sheet_name} · block {block.index} · {block.orientation.value} · "
            f"extraction {'row' if block.orientation.value == 'row-wise' else 'column'} "
            f"{block.header_ref} · selector {block.info_ref}"
        )
        self._line(
            "Fields resolved by label: "
            + " · ".join(f"{k} → {v}" for k, v in block.address_map.items())
        )
        declared = " · ".join(
            f"{a.name} = {a.value}{' (hypothesis)' if a.is_hypothesis else ''}"
            for a in block.attributes.values()
        )
        self._line(f"Attributes: {declared}" if declared else "Attributes: none declared")
        self._line(
            f"Records: {block.candidates} candidates, {len(block.records)} extracted, "
            f"{block.excluded} excluded (selector empty)"
        )
        self._line(f"Confidence: {block.confidence.value} (worst of all hypotheses)")
        if block.unextracted:
            self._line(
                "Contained data but not extracted, not declared in sheet 00: "
                + ", ".join(block.unextracted)
            )
        self.row += 1

        if block.hypotheses:
            self._put(FIRST_COL, "Assumptions & hypotheses", head_f)
            self.row += 1
            for h in block.hypotheses:
                self._put(FIRST_COL, h.id, body_f)
                self._put(FIRST_COL + 1, h.attribute, body_f)
                self._put(FIRST_COL + 2, "—" if h.value is None else str(h.value), body_f)
                self._put(FIRST_COL + 3, h.confidence.value, body_f)
                self._put(FIRST_COL + 4, h.status, body_f)
                self._put(FIRST_COL + 5, h.note, note_f)
                self.row += 1
            self.row += 1

        columns = [block.provenance_label, *ds.headers]
        if block.orientation.value == "transposed":
            columns.append("Confidence")

        for i, label in enumerate(columns):
            self._put(FIRST_COL + i, label, head_f, fill=step1_fill)
        self.row += 1

        first_data = self.row
        for record in block.records:
            self._put(FIRST_COL, record.source_ref, body_f)
            for i, label in enumerate(ds.headers, start=1):
                self._put(FIRST_COL + i, record.values.get(label), body_f,
                          _number_format(block, label))
            if block.orientation.value == "transposed":
                self._put(FIRST_COL + len(ds.headers) + 1,
                          record.confidence.value if record.confidence else "", note_f)
            self.row += 1
        last_data = self.row - 1

        self._write_controls(block, ds.headers, first_data, last_data, block.totals())
        self._write_excluded_note(block)

    def _write_controls(self, block, headers, first_data, last_data, totals):
        self._put(FIRST_COL, f"Control  (n = {len(block.records)})", ctrl_f, fill=ctrl_fill)
        control_row = self.row
        for i, label in enumerate(headers, start=1):
            if label == block.dataset.key_field:
                continue
            letter = col_letter(FIRST_COL + i)
            self._formula(FIRST_COL + i, f"=SUM({letter}{first_data}:{letter}{last_data})",
                          totals.get(label), _number_format(block, label), ctrl_f, ctrl_fill)
        self.row += 1

        self._put(FIRST_COL, "Expected (computed by tool)", note_f)
        expected_row = self.row
        for i, label in enumerate(headers, start=1):
            if label == block.dataset.key_field:
                continue
            self._put(FIRST_COL + i, totals.get(label), note_f, _number_format(block, label))
        self.row += 1

        self._put(FIRST_COL, "Check", ctrl_f)
        for i, label in enumerate(headers, start=1):
            if label == block.dataset.key_field:
                continue
            letter = col_letter(FIRST_COL + i)
            self._formula(
                FIRST_COL + i,
                f'=IF(ABS({letter}{control_row}-{letter}{expected_row})<=0.5,"OK","MISMATCH")',
                "OK", None, ctrl_f,
            )
        self.row += 1

    def _write_excluded_note(self, block: Block):
        if not block.excluded_records:
            return
        totals = {
            m: sum(r.values[m] for r in block.excluded_records
                   if isinstance(r.values.get(m), (int, float)))
            for m in block.dataset.measures
        }
        self._put(FIRST_COL, "Excluded records, measure totals", note_f)
        for i, label in enumerate(block.dataset.headers, start=1):
            if label == block.dataset.key_field:
                continue
            self._put(FIRST_COL + i, totals.get(label), note_f, _number_format(block, label))
        self.row += 1
        self._line(
            "Shown as an independent control: an excluded total row should equal the "
            "extracted total above. No assumption is made about what the excluded records are."
        )

    # ────────────────────────────────────────────────────────── step 2
    def write_step2(self, result: Step2Result) -> None:
        block, ds = result.block, result.block.dataset
        self.row += GAP
        self._put(FIRST_COL, anchor_tag(ds.key, "STEP2"), anchor_f)
        self.row += 1
        self._put(FIRST_COL, "STEP 2 — NORMALISED", title_f, fill=step2_fill)
        self.row += 1

        for note in result.notes:
            self._line(f"· {note}")
        self._line(f"Confidence carried from step 1: {block.confidence.value}")
        self.row += 1

        columns = [block.provenance_label, *result.columns]
        for i, label in enumerate(columns):
            self._put(FIRST_COL + i, label, head_f, fill=step2_fill)
        self.row += 1

        calc_names = [c.name for c in result.spec.calculations]
        col_of = {label: FIRST_COL + 1 + i for i, label in enumerate(result.columns)}

        first_data = self.row
        for n, record in enumerate(result.records):
            self._put(FIRST_COL, record.source_ref, body_f)
            for label in result.spec.column_order:
                self._put(col_of[label], record.values.get(label), body_f,
                          _number_format(block, label))
            for calc in result.spec.calculations:
                expected = result.computed[calc.name][n]
                formula = calc.expression
                for fld in result.spec.column_order:
                    formula = formula.replace(
                        "{" + fld + "}", f"{col_letter(col_of[fld])}{self.row}"
                    )
                self._formula(col_of[calc.name], f"=IFERROR({formula},\"\")",
                              expected, calc.number_format, body_f)
            self.row += 1
        last_data = self.row - 1

        measures = [m for m in ds.measures]
        self._put(FIRST_COL, f"Control  (n = {len(result.records)})", ctrl_f, fill=ctrl_fill)
        control_row = self.row
        totals = result.totals()
        for label in measures:
            letter = col_letter(col_of[label])
            self._formula(col_of[label], f"=SUM({letter}{first_data}:{letter}{last_data})",
                          totals.get(label), _number_format(block, label), ctrl_f, ctrl_fill)
        self.row += 1

        self._put(FIRST_COL, "Tie-back to step 1 (must be unchanged)", note_f)
        for label in measures:
            self._put(col_of[label], block.totals().get(label), note_f,
                      _number_format(block, label))
        self.row += 1

        self._put(FIRST_COL, "Check", ctrl_f)
        tie_row = self.row - 1
        for label in measures:
            letter = col_letter(col_of[label])
            self._formula(
                col_of[label],
                f'=IF(ABS({letter}{control_row}-{letter}{tie_row})<=0.5,"OK","MISMATCH")',
                "OK", None, ctrl_f,
            )
        self.row += 1
        if calc_names:
            self._line(
                f"{', '.join(calc_names)} is value-adding, so no tie-back applies; "
                "it is written as a live formula over the columns above."
            )


def write_blocks(ws, block: Block, result: Step2Result) -> list[tuple[str, str, float]]:
    """Write step 1 and step 2 beneath the last non-empty row — spec §9.1."""
    from .extract import last_non_empty_row

    clear_generated(ws)
    start = last_non_empty_row(ws) + 1 + GAP
    writer = BlockWriter(ws, start)
    writer.write_step1(block)
    writer.write_step2(result)

    for col in range(FIRST_COL, FIRST_COL + 8):
        letter = col_letter(col)
        if ws.column_dimensions[letter].width in (None, 0):
            ws.column_dimensions[letter].width = 16
    return writer.formula_values
