"""Command line entry point: `packing-spec run ...` and `packing-spec demo`."""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from .dims import TOLERANCE_MM
from .pipeline import RunConfig, run, summary_lines

REPO_ROOT = Path(__file__).resolve().parents[2]
SAMPLE_DIR = REPO_ROOT / "examples" / "sample_data"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="packing-spec",
        description="Compare ERP packing data, the packing instructions sent to the vendor "
                    "and the vendor packing list, SKU by SKU.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    r = sub.add_parser("run", help="Run on an input folder or on explicit files.")
    r.add_argument("--input-dir", type=Path, help="Folder with the input workbooks "
                   "(recognized by content).")
    r.add_argument("--catalog", type=Path, required=True,
                   help="SKU catalog workbook (internal / parent / pack size / marketplace).")
    r.add_argument("--erp", type=Path, help="ERP product export (skips folder discovery).")
    r.add_argument("--packing-list", type=Path, help="Vendor packing list.")
    r.add_argument("--instructions", type=Path, action="append", default=[],
                   help="Packing-instructions file. Repeatable; ORDER IS PRIORITY "
                        "(first wins, the rest only fill gaps).")
    r.add_argument("--pl-sheet", help="Force the packing-list tab.")
    r.add_argument("--vendor", help="Vendor name for output file names.")
    r.add_argument("--po", help="PO label for output file names (default: from file names).")
    r.add_argument("--tolerance-mm", type=float, default=TOLERANCE_MM,
                   help=f"Max per-side carton difference treated as equal (default "
                        f"{TOLERANCE_MM:g}).")
    r.add_argument("--out", type=Path, default=Path("output"), help="Output folder.")

    d = sub.add_parser("demo", help="Run end-to-end on the bundled synthetic sample data.")
    d.add_argument("--out", type=Path, default=Path("output"), help="Output folder.")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(message)s")

    if args.command == "demo":
        if not (SAMPLE_DIR / "input").exists():
            print("Sample data not found. Run: python examples/generate_sample_data.py")
            return 1
        cfg = RunConfig(catalog=SAMPLE_DIR / "reference" / "sku_catalog.xlsx",
                        input_dir=SAMPLE_DIR / "input", out_dir=args.out,
                        vendor="Northwind")
    else:
        if args.instructions and not args.erp:
            print("--instructions requires --erp (explicit mode); otherwise use --input-dir.")
            return 2
        cfg = RunConfig(catalog=args.catalog, out_dir=args.out, input_dir=args.input_dir,
                        erp=args.erp, packing_list=args.packing_list,
                        instructions=args.instructions, pl_sheet=args.pl_sheet,
                        vendor=args.vendor, po=args.po, tolerance_mm=args.tolerance_mm)

    result = run(cfg)
    print("\n".join(summary_lines(result)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
