"""CLI entry point for PowerPoint Excel report automation."""

import argparse
import glob
import logging
import os
import sys
import time

import yaml

from ppt_automation.session import Session
from ppt_automation import linker, table_updater, delta_updater, color_coder, chart_updater
from ppt_automation.shape_finder import build_presentation_inventory


def load_config(config_path: str) -> dict:
    """Load configuration from a YAML file."""
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def resolve_paths(patterns: list[str]) -> list[str]:
    """Resolve glob patterns to absolute file paths."""
    paths = []
    for pattern in patterns:
        expanded = glob.glob(pattern)
        if expanded:
            paths.extend(os.path.abspath(p) for p in expanded)
        else:
            # Treat as literal path
            paths.append(os.path.abspath(pattern))
    return paths


def parse_pair(pair_str: str) -> tuple[str, str]:
    """Parse a 'pptx:xlsx' pair string into (pptx_path, excel_path)."""
    if ":" not in pair_str:
        print(f"Invalid pair format: '{pair_str}'. Expected 'file.pptx:data.xlsx'")
        sys.exit(1)
    # Split on last ':' to handle Windows drive letters like C:\path
    # Find the colon that separates pptx from xlsx (not a drive letter colon)
    # Strategy: split on ':', rejoin if we accidentally split a drive letter
    parts = pair_str.split(":")
    if len(parts) == 3:
        # e.g. C:\file.pptx:C:\data.xlsx -> impossible, both have drive letters
        # More likely: file.pptx:C:\data.xlsx or C:\file.pptx:data.xlsx
        # Try: first part is just a drive letter -> rejoin
        if len(parts[0]) == 1 and parts[0].isalpha():
            # "C:\file.pptx:data.xlsx" -> pptx="C:\file.pptx", excel="data.xlsx"
            pptx = f"{parts[0]}:{parts[1]}"
            excel = parts[2]
        else:
            # "file.pptx:C:\data.xlsx" -> pptx="file.pptx", excel="C:\data.xlsx"
            pptx = parts[0]
            excel = f"{parts[1]}:{parts[2]}"
    elif len(parts) == 4:
        # "C:\file.pptx:C:\data.xlsx"
        pptx = f"{parts[0]}:{parts[1]}"
        excel = f"{parts[2]}:{parts[3]}"
    elif len(parts) == 2:
        pptx, excel = parts
    else:
        print(f"Invalid pair format: '{pair_str}'")
        sys.exit(1)
    return os.path.abspath(pptx), os.path.abspath(excel)


def process_presentation(
    pptx_path: str,
    excel_path: str,
    config: dict,
    options: argparse.Namespace,
) -> dict:
    """Process a single presentation through the full pipeline.

    Returns a dict with counts: links, tables, deltas, colors, charts.
    """
    results = {"links": 0, "tables": 0, "deltas": 0, "colors": 0, "charts": 0}

    with Session(pptx_path, excel_path) as session:
        # Build shape inventory ONCE — all steps use O(1) lookups from this
        inventory = build_presentation_inventory(session.presentation)

        if not options.skip_links:
            results["links"] = linker.update_links(
                session, excel_path, config, inventory=inventory
            )

        results["tables"] = table_updater.update_tables(
            session, config, inventory=inventory
        )

        if not options.skip_deltas:
            results["deltas"] = delta_updater.update_deltas(
                session, config, inventory=inventory
            )

        if not options.skip_coloring:
            results["colors"] = color_coder.apply_color_coding(
                session, config, inventory=inventory
            )

        if not options.skip_charts:
            results["charts"] = chart_updater.update_charts(
                session, excel_path, inventory=inventory
            )

        session.save()

    return results


def _run_pairs(pairs: list[tuple[str, str]], config: dict, args: argparse.Namespace):
    """Run the pipeline for a list of (pptx_path, excel_path) pairs."""
    grand_total = {"links": 0, "tables": 0, "deltas": 0, "colors": 0, "charts": 0}
    t_start = time.perf_counter()
    processed = 0

    for pptx_path, excel_path in pairs:
        if not os.path.exists(pptx_path):
            print(f"PPT not found, skipping: {pptx_path}")
            continue
        if not os.path.exists(excel_path):
            print(f"Excel not found, skipping: {excel_path}")
            continue

        print(f"Processing: {os.path.basename(pptx_path)} <- {os.path.basename(excel_path)}")
        t_file = time.perf_counter()

        results = process_presentation(pptx_path, excel_path, config, args)

        elapsed = time.perf_counter() - t_file
        print(
            f"  Done in {elapsed:.2f}s — "
            f"{results['links']} links, "
            f"{results['tables']} tables, "
            f"{results['deltas']} deltas, "
            f"{results['colors']} colored, "
            f"{results['charts']} charts"
        )

        for key in grand_total:
            grand_total[key] += results[key]
        processed += 1

    total_elapsed = time.perf_counter() - t_start
    print(
        f"\nAll done! {processed} file(s) in {total_elapsed:.2f}s\n"
        f"  {grand_total['links']} link(s) updated\n"
        f"  {grand_total['tables']} table(s) refreshed\n"
        f"  {grand_total['deltas']} delta(s) updated\n"
        f"  {grand_total['colors']} table(s) color-coded\n"
        f"  {grand_total['charts']} chart(s) updated"
    )


def main():
    parser = argparse.ArgumentParser(
        description="PowerPoint Excel report automation — Python rewrite",
        epilog=(
            "Examples:\n"
            "  %(prog)s report.pptx --excel data.xlsx\n"
            "  %(prog)s report.pptx                        (file picker opens)\n"
            '  %(prog)s --pair "us.pptx:us.xlsx" --pair "mx.pptx:mx.xlsx"\n'
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "presentations",
        nargs="*",
        help="One or more .pptx file paths (supports glob patterns). Used with --excel.",
    )
    parser.add_argument(
        "--excel", "-e",
        default=None,
        help="Path to the Excel data file. If omitted, a file dialog will open.",
    )
    parser.add_argument(
        "--pair", "-p",
        action="append",
        default=None,
        metavar="PPT:XLSX",
        help="A pptx:xlsx pair. Can be repeated for batch processing multiple pairs.",
    )
    parser.add_argument(
        "--config", "-c",
        default=None,
        help="Path to config.yaml (default: config.yaml next to this script)",
    )
    parser.add_argument("--skip-links", action="store_true", help="Skip Step 1a (re-link OLE)")
    parser.add_argument("--skip-deltas", action="store_true", help="Skip Step 1c (delta arrows)")
    parser.add_argument("--skip-coloring", action="store_true", help="Skip Step 1d (_ccst coloring)")
    parser.add_argument("--skip-charts", action="store_true", help="Skip Step 2 (chart links)")
    parser.add_argument("--verbose", "-v", action="store_true", help="Enable debug logging")

    args = parser.parse_args()

    # Logging
    level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )

    # Config
    if args.config:
        config_path = args.config
    else:
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.yaml")
    config = load_config(config_path)

    # --- Mode 1: --pair for explicit pptx:xlsx pairs ---
    if args.pair:
        pairs = [parse_pair(p) for p in args.pair]
        _run_pairs(pairs, config, args)
        return

    # --- Mode 2: presentations + --excel (or file picker) ---
    if not args.presentations:
        parser.error("Provide presentation file(s) or use --pair for batch pairs.")

    excel_path = args.excel
    if not excel_path:
        from ppt_automation.file_picker import pick_excel_file
        excel_path = pick_excel_file()
        if not excel_path:
            print("No Excel file selected. Exiting.")
            sys.exit(1)
    excel_path = os.path.abspath(excel_path)

    if not os.path.exists(excel_path):
        print(f"Excel file not found: {excel_path}")
        sys.exit(1)

    pptx_files = resolve_paths(args.presentations)
    if not pptx_files:
        print("No presentation files found.")
        sys.exit(1)

    pairs = [(p, excel_path) for p in pptx_files]
    _run_pairs(pairs, config, args)


if __name__ == "__main__":
    main()
