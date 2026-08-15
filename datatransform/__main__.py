"""CLI:  python -m datatransform Intake_v1.xlsx [-o output.xlsx]"""

import argparse
import sys

from .model import ExtractionError
from .runner import run


def main(argv=None) -> int:
    p = argparse.ArgumentParser(prog="datatransform", description=__doc__)
    p.add_argument("source", help="workbook to process (never modified)")
    p.add_argument("-o", "--output", default=None, help="output workbook")
    p.add_argument("--log-dir", default="logs")
    args = p.parse_args(argv)

    try:
        report = run(args.source, args.output, args.log_dir)
    except ExtractionError as exc:
        print(f"FAILED: {exc}", file=sys.stderr)
        return 2

    for o in report.outcomes:
        print(f"{o.sheet:<28} {o.status:<10} {o.detail}")
    print(f"\noutput : {report.output}")
    print(f"sha256 : {report.source_sha256[:16]}…  ({report.injected} formula values cached)")
    return 0 if report.ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
