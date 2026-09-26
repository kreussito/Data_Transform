"""CLI:  python -m datatransform Intake_v1.xlsx [-o output.xlsx]"""

from __future__ import annotations

import argparse
import sys

from .console import safe, use_utf8
from .constants import VERSION
from .model import ExtractionError
from .runner import run

SHEET_WIDTH = 28
STATUS_WIDTH = 10
SHA_SHOWN = 16


def main(argv=None) -> int:
    # Before anything is printed: the summary contains ⟦…⟧ and →, and cp1252 — still the
    # default for redirected output on Windows — has neither.
    use_utf8()

    p = argparse.ArgumentParser(prog="datatransform", description=__doc__)
    p.add_argument("source", help="workbook to process (never modified)")
    p.add_argument("-o", "--output", default=None, help="output workbook")
    p.add_argument("--log-dir", default="logs")
    p.add_argument("--version", action="version", version=f"datatransform {VERSION}")
    args = p.parse_args(argv)

    try:
        report = run(args.source, args.output, args.log_dir)
    except ExtractionError as exc:
        print(safe(f"FAILED: {exc}", sys.stderr), file=sys.stderr)
        return 2

    for o in report.outcomes:
        print(safe(f"{o.sheet:<{SHEET_WIDTH}} {o.status:<{STATUS_WIDTH}} {o.detail}"))
    print(safe(f"\noutput : {report.output}"))
    print(safe(f"sha256 : {report.source_sha256[:SHA_SHOWN]}…  "
               f"({report.injected} formula values cached)"))
    return 0 if report.ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
