"""The shared vocabulary — datatransform/constants.py.

These are not settings, so there is nothing here about reading them from a file. What
there is: proof that the values several modules depend on are defined once, and that the
sets which have to be complete actually are. A constant nobody checks is a constant that
drifts.
"""

from __future__ import annotations

import ast
import pathlib

import pytest

from datatransform import constants as K
from datatransform.bridge import DEFAULT_VIEW_THRESHOLD, VIEW_ATTRIBUTE
from datatransform.growth import DEFAULT_THRESHOLD as RATE_DEFAULT
from datatransform.growth import EXPOSURE_FIELD, THRESHOLD_ATTRIBUTE as RATE_ATTR
from datatransform.lossshare import DEFAULT_THRESHOLD as LOSS_DEFAULT
from datatransform.lossshare import THRESHOLD_ATTRIBUTE as LOSS_ATTR

PACKAGE = pathlib.Path(__file__).resolve().parents[1] / "datatransform"


class _Globals:
    def __init__(self, **values):
        self.globals = values


# ─────────────────────────────────────── the aliases resolve to one value

def test_each_module_alias_points_at_the_shared_constant():
    """Modules keep readable local names; the value behind them is defined once."""
    assert EXPOSURE_FIELD is K.F_TOTAL
    assert (RATE_ATTR, RATE_DEFAULT) == (K.RATE_CHANGE_WARNING, K.RATE_CHANGE_DEFAULT)
    assert (LOSS_ATTR, LOSS_DEFAULT) == (K.LOSS_SHARE_WARNING, K.LOSS_SHARE_DEFAULT)
    assert (VIEW_ATTRIBUTE, DEFAULT_VIEW_THRESHOLD) == (K.SPLIT_VIEW_WARNING,
                                                        K.SPLIT_VIEW_DEFAULT)


def test_the_three_thresholds_are_three_distinct_attributes():
    """They happen to share a default; they must never share a name."""
    names = {K.LOSS_SHARE_WARNING, K.RATE_CHANGE_WARNING, K.SPLIT_VIEW_WARNING}
    assert len(names) == 3


def test_the_version_is_stated_once():
    from datatransform import __version__
    from datatransform.runner import TOOL_VERSION

    assert __version__ == TOOL_VERSION == K.VERSION


# ─────────────────────────────────────────────── sets that must be complete

def test_every_reporting_level_has_a_readable_name():
    """LEVEL_NAMES is what the underwriter reads in the finding, so a level missing
    from it would print as a bare identifier."""
    levels = {v for name, v in vars(K).items()
              if name.startswith("LEVEL_") and isinstance(v, str)}
    assert levels == set(K.LEVEL_NAMES)


def test_the_structural_markers_are_the_four_of_section_4():
    from datatransform.markers import STRUCTURAL

    assert STRUCTURAL is K.STRUCTURAL_MARKERS
    assert set(K.STRUCTURAL_MARKERS) == {K.M_HEADER, K.M_INFO, K.M_TRANSPOSE,
                                         K.M_DATASET}
    assert "Section" not in K.STRUCTURAL_MARKERS      # an attribute, not structure


def test_the_role_groups_are_drawn_from_the_declared_roles():
    roles = {v for name, v in vars(K).items()
             if name.startswith("ROLE_") and isinstance(v, str)}
    assert set(K.LOSS_ROLES) <= roles
    assert set(K.EXPOSURE_ROLES) <= roles
    assert not set(K.LOSS_ROLES) & set(K.EXPOSURE_ROLES)


# ────────────────────────────────────────────────── one reader, three uses

@pytest.mark.parametrize("declared,expected", [
    ("20%", 0.20), ("20", 0.20), (0.2, 0.20), (15, 0.15), ("0.15", 0.15),
    ("", 0.99), (None, 0.99), ("nonsense", 0.99),
])
def test_a_declared_share_reads_the_same_way_everywhere(declared, expected):
    """20%, 20 and 0.2 all mean twenty per cent — a share above 1 can only have been
    meant as a percentage."""
    assert K.read_share(_Globals(T=declared), "T", 0.99) == pytest.approx(expected)


def test_an_undeclared_threshold_falls_back_to_the_tool_s_opinion():
    assert K.read_share(_Globals(), K.LOSS_SHARE_WARNING,
                        K.LOSS_SHARE_DEFAULT) == K.LOSS_SHARE_DEFAULT


# ───────────────────────────────────── the duplication that started this

def test_no_module_redefines_a_shared_literal():
    """The failure this file exists to prevent: two modules each writing out the same
    value, and one of them being edited later."""
    shared = {K.F_TOTAL, K.LOSS_SHARE_WARNING, K.RATE_CHANGE_WARNING,
              K.SPLIT_VIEW_WARNING, K.CONTROL_OK, K.CONTROL_MISMATCH,
              K.NOT_APPLICABLE, K.M_HEADER}
    offenders = []
    for path in sorted(PACKAGE.glob("*.py")):
        if path.name == "constants.py":
            continue
        for node in ast.walk(ast.parse(path.read_text(encoding="utf-8"))):
            if isinstance(node, ast.Constant) and node.value in shared:
                offenders.append(f"{path.name}:{node.lineno} {node.value!r}")
    assert not offenders, "write these through datatransform.constants: " + \
        ", ".join(offenders)


def test_the_percent_regex_is_written_once():
    written = [p.name for p in PACKAGE.glob("*.py")
               if "re.compile(r\"^\\s*([0-9.,]+)\\s*%" in p.read_text(encoding="utf-8")]
    assert written == ["constants.py"]
