"""Step-2 mechanics, held in code rather than in sheet 00.

Specification_v1.md §3: sheet 00 carries facts a human knows and a machine cannot
infer. Sort order, target column order and derived measures are *mechanics*, so they
live here, versioned with the tool.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .model import FieldType

SPEC_VERSION = "1"

# How each declared field is read — spec §8.4. Declared per field rather than inferred
# from position, so that a non-numeric field sitting where a measure would (03. Large
# Losses: Loss Date, Claim Reference) is never summed.
FIELD_TYPES: dict[str, dict[str, FieldType]] = {
    "01 History": {
        "Year": FieldType.TEXT,
        "Premium": FieldType.NUMBER,
        "Incurred Losses": FieldType.NUMBER,
    },
}
FIELD_TYPES["01 History_T"] = FIELD_TYPES["01 History"]


def field_types_for(dataset_key: str, headers) -> dict[str, FieldType]:
    """Declared types, defaulting to NUMBER for anything not yet specified."""
    declared = FIELD_TYPES.get(dataset_key, {})
    return {h: declared.get(h, FieldType.NUMBER) for h in headers}


@dataclass(frozen=True)
class Calculation:
    """A derived measure. ``expression`` names fields in braces."""

    name: str
    expression: str
    number_format: str = "0.0%"
    guard_zero: str | None = None      # field that must not be zero


@dataclass(frozen=True)
class Step2Spec:
    sort_by: tuple[str, ...]
    ascending: bool = True
    column_order: tuple[str, ...] = ()
    calculations: tuple[Calculation, ...] = field(default_factory=tuple)

    def output_columns(self) -> tuple[str, ...]:
        return tuple(self.column_order) + tuple(c.name for c in self.calculations)


STEP2: dict[str, Step2Spec] = {
    "01 History": Step2Spec(
        sort_by=("Year",),
        ascending=True,
        column_order=("Year", "Premium", "Incurred Losses"),
        calculations=(
            Calculation(
                name="Loss Ratio %",
                expression="{Incurred Losses}/{Premium}",
                number_format="0.0%",
                guard_zero="Premium",
            ),
        ),
    ),
}

# The transposed reference sheet is the same dataset in a different orientation.
STEP2["01 History_T"] = STEP2["01 History"]


def step2_for(dataset_key: str) -> Step2Spec | None:
    return STEP2.get(dataset_key)
