"""The words and numbers more than one module depends on.

**Not configuration.** The workbook is where this tool is configured: sheet ``00``
declares the datasets, the vocabulary, the sections, the zones, the axes, the split
ratios and the three warning thresholds, and §3 says why — those are facts a human knows
and a machine cannot infer, so they belong where the human is. Adding a second
configuration surface beside sheet ``00`` would raise a question with no good answer:
which one wins.

What lives here is the opposite kind of value — the ones that must **never** vary,
because they are the code's own vocabulary. ``"occupancy"`` is not a setting; it is how
two modules agree on what they are talking about. Renaming it in a config file would not
change the tool's behaviour, it would break it.

The reason to collect them is narrower and more practical: several were declared
*independently in two places*. ``"Total"`` was ``TOTAL`` in one module and
``EXPOSURE_FIELD`` in another; the percentage regex was written out twice; two modules
each defined ``DEFAULT_THRESHOLD = 0.20`` for different thresholds that happened to share
a value. None of that was wrong today, and all of it was one careless edit away from
being wrong tomorrow. Modules keep their own readable aliases — ``EXPOSURE_FIELD`` still
reads better than ``F_TOTAL`` inside the growth block — but the value is defined once.
"""

from __future__ import annotations

import re

VERSION = "0.1.0"

# ── dataset roles ─────────────────────────────────────────────────────────────
# The leading number of a dataset key. A section declares which it expects (§2.3), so
# an absent role is structure rather than a gap.
ROLE_HISTORY = "01"
ROLE_EPI = "02"
ROLE_LARGE = "03"
ROLE_CAT = "04"
ROLE_PROFILE = "05"
ROLE_EQ_AGGS = "06"
ROLE_WIND_AGGS = "07"
ROLE_SPLITS = "08"

LOSS_ROLES = (ROLE_LARGE, ROLE_CAT)              # §10.3 — what is summed
EXPOSURE_ROLES = (ROLE_EQ_AGGS, ROLE_WIND_AGGS)  # §10.4 — what carries an aggregate

# ── field names the code reaches for by name ──────────────────────────────────
# Everything else is resolved through sheet 00. These few are named here because the
# arithmetic itself refers to them: a growth block has to know which column is exposure.
F_YEAR = "Year"
F_PREMIUM = "Premium"
F_INCURRED = "Incurred Losses"
F_EPI = "EPI"
F_LOSS_AMOUNT = "Loss amount"
F_TOTAL = "Total"
F_ZONE = "Zone"
F_CATEGORY = "Category"
F_EXPOSURE = "Exposure"

# ── attributes ────────────────────────────────────────────────────────────────
A_SECTION = "Section"
A_CURRENCY = "Currency"
A_SCALE = "Scale"
A_PERIOD = "Period"
A_AS_AT = "As at"
A_LOSS_BASIS = "Loss basis"
A_SHARE_BASIS = "Share basis"
A_PREMIUM_BASIS = "Premium basis"

# ── the three declared thresholds ─────────────────────────────────────────────
# Each is an ⟦GLOBAL⟧ attribute with a fallback. The fallback is the tool's opinion; the
# attribute is the underwriter's, and the underwriter's wins.
LOSS_SHARE_WARNING = "Loss share warning"
LOSS_SHARE_DEFAULT = 0.20

RATE_CHANGE_WARNING = "Rate change warning"
RATE_CHANGE_DEFAULT = 0.20

SPLIT_VIEW_WARNING = "Split view warning"
SPLIT_VIEW_DEFAULT = 0.02

PERCENT = re.compile(r"^\s*([0-9.,]+)\s*%\s*$")


def read_share(nomenclature, attribute: str, default: float) -> float:
    """One ⟦GLOBAL⟧ share, read the same way wherever it is used.

    ``20%``, ``20`` and ``0.2`` all mean twenty per cent: a share above 1 can only have
    been meant as a percentage. Shared by all three thresholds so they cannot drift into
    reading their own numbers differently.
    """
    raw = (getattr(nomenclature, "globals", None) or {}).get(attribute)
    text = "" if raw is None else str(raw).strip()
    if not text:
        return default
    percent = PERCENT.match(text)
    try:
        value = float((percent.group(1) if percent else text).replace(",", ""))
    except ValueError:
        return default
    return value / 100.0 if (percent or value > 1) else value


# ── tolerances ────────────────────────────────────────────────────────────────
# Three different questions, so three different numbers rather than one shared "epsilon"
# that would quietly mean something different in each place.
IDENTITY_TOLERANCE = 0.5      # a Total against its parts: rounding, not disagreement
LOSS_TOLERANCE = 1.0          # §10.3, on figures already scaled to thousands
SHARE_CLOSURE = 0.005         # ratios that must sum to 100%
IPF_ROUNDS = 60               # fitting to two margins: rounds, and when to stop
IPF_TOLERANCE = 1e-9
UNRANKED_PERIOD = 9999        # a suffix ⟦PERIOD ORDER⟧ does not rank sorts last

PERCENT_SCALE = 100.0

# ── marker keywords ───────────────────────────────────────────────────────────
# The four structural markers of column A (§4). Everything else in column A is an
# attribute; these say where a block *is*, so extract, markers and runner all have to
# agree on them to the letter.
M_HEADER = "Header"
M_INFO = "Info"
M_TRANSPOSE = "Transpose"
M_DATASET = "Dataset"
STRUCTURAL_MARKERS = (M_HEADER, M_INFO, M_TRANSPOSE, M_DATASET)

A_OCCURRENCE_FROM = "Occurrence year from"
NOT_AVAILABLE = "n/a"          # a figure that could not be computed, in a written line

# ── the reporting levels of §2.6 ──────────────────────────────────────────────
# How the bridge and step 2 agree on what a block reported. Bare strings on both sides
# of that boundary is exactly the coupling that goes wrong in silence.
LEVEL_REPORTED = "reported"     # the finished target grid
LEVEL_BOTH = "both"             # occupancy and cover, as two vectors
LEVEL_OCCUPANCY = "occupancy"
LEVEL_COVER = "cover"
LEVEL_SEGMENT = "segment"       # Projects / Renewables — not a target axis at all
LEVEL_TOTAL = "total"

LEVEL_NAMES = {
    LEVEL_REPORTED: "reported grid",
    LEVEL_OCCUPANCY: "occupancy split",
    LEVEL_COVER: "cover split",
    LEVEL_BOTH: "two reported margins",
    LEVEL_SEGMENT: "Projects/Renewables segmentation",
    LEVEL_TOTAL: "single zone total",
}

# ── outcome words ─────────────────────────────────────────────────────────────
# A rule has four outcomes, not two (§10.1): "not applicable" and "skipped" are both
# *not evaluated*, and the difference between them is the whole point.
PASSED = "passed"
FAILED = "failed"
NOT_APPLICABLE = "not applicable"
SKIPPED = "skipped"

# The control check writes these; the group checks deliberately do not, so that one word
# never means two things on one sheet (§10.3).
CONTROL_OK = "OK"
CONTROL_MISMATCH = "MISMATCH"

WITHIN_LIMITS = "within limits"
WARNING = "WARNING"
NO_BASIS = "no basis"

# ── sheet outcomes ────────────────────────────────────────────────────────────
PROCESSED = "processed"
ERROR = "error"

# ── number formats ────────────────────────────────────────────────────────────
FMT_AMOUNT = "#,##0"
FMT_PERCENT = "0.0%"
FMT_PERCENT_ROUND = "0%"
FMT_TEXT = "@"
FMT_INTEGER = "0"
FMT_DATE = "yyyy-mm-dd"
FMT_LOG_AMOUNT = ",.0f"       # the same figure in a log line

# ── log layout ────────────────────────────────────────────────────────────────
LOG_RULE_WIDTH = 78
NOTE_EXAMPLES = 3             # how many worked examples a note shows before "…"
FINDING_EXAMPLES = 5          # how many disagreeing records a finding names
