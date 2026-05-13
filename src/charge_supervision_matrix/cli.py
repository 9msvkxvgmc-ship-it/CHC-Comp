"""
CLI entry point for charge-supervision-matrix.

Usage examples
--------------
    charge-supervision-matrix charges.xlsx
    charge-supervision-matrix charges.xlsx -o report.xlsx
    charge-supervision-matrix charges.xlsx --exclude "Hogan, Jennifer" --exclude "Smith, Bob"
    charge-supervision-matrix charges.xlsx --add-app "Tol, Margie" --reclassify-as-md "Eastburg MD, Luke"
    charge-supervision-matrix charges.xlsx --wrvu-file rvu26a.csv --location "Mercy Health"
"""

import argparse
import sys
from .config import Config
from .runner import run
from .suggest import print_report


def main():
    parser = argparse.ArgumentParser(
        prog="charge-supervision-matrix",
        description="Generate Charge Summary and Supervision Matrix from an iKnowMed All Signed Charges export.",
    )
    parser.add_argument("input", help="Path to the source xlsx file")
    parser.add_argument(
        "--suggest",
        action="store_true",
        help="Analyze signers and print recommended flags without generating a report",
    )
    parser.add_argument("-o", "--output", default=None, help="Output xlsx path (default: <input>_report.xlsx)")
    parser.add_argument(
        "--exclude",
        action="append",
        default=[],
        metavar="NAME",
        help="Exclude signer from all output (repeatable)",
    )
    parser.add_argument(
        "--reclassify-as-md",
        action="append",
        default=[],
        metavar="NAME",
        help="Treat this signer as a supervising MD (repeatable)",
    )
    parser.add_argument(
        "--add-app",
        action="append",
        default=[],
        metavar="NAME",
        help="Add this name to the APP list even without PA/NP credentials (repeatable)",
    )
    parser.add_argument(
        "--omit",
        action="append",
        default=[],
        metavar="NAME",
        help="Completely omit this person from output (repeatable)",
    )
    parser.add_argument(
        "--report-type",
        choices=["inpatient", "outpatient"],
        default="inpatient",
        help="Report type: inpatient (grouped by Supervising MD) or outpatient (grouped by Order MD) [default: inpatient]",
    )
    parser.add_argument(
        "--wrvu-file",
        default=None,
        metavar="PATH",
        help="Path to CMS RVU CSV file (optional; built-in table used otherwise)",
    )
    parser.add_argument(
        "--location",
        default=None,
        metavar="FILTER",
        help="Only include charges from locations matching this substring",
    )

    args = parser.parse_args()

    if args.suggest:
        try:
            print_report(args.input, report_type=args.report_type)
        except Exception as exc:
            print(f"Error: {exc}", file=sys.stderr)
            sys.exit(1)
        return

    config = Config(
        excluded_signers=args.exclude,
        reclassify_as_supervising_md=args.reclassify_as_md,
        add_to_app_list=args.add_app,
        omit=args.omit,
        wrvu_file=args.wrvu_file,
        location_filter=args.location,
    )

    try:
        out = run(input_path=args.input, output_path=args.output, config=config, report_type=args.report_type)
        print(f"Written: {out}")
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
