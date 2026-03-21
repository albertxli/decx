"""CLI entry point for decx — PowerPoint Excel report automation."""

import argparse
import glob
import logging
import os
import shutil
import sys
import time
from collections import Counter

from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn, TimeElapsedColumn
from rich.table import Table

from decx import __version__
from decx.config import get_config, DEFAULT_CONFIG
from decx.session import Session
from decx import linker, table_updater, delta_updater, color_coder, chart_updater
from decx.shape_finder import build_presentation_inventory, find_template_shape

console = Console()

MSO_LINKED_OLE_OBJECT = 10


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


def resolve_output_path(
    pptx_path: str, output: str | None, is_batch: bool, pair_count: int
) -> str:
    """Determine the actual pptx path to process, copying if output specified.

    Returns the path to process (may be a copy of the original).
    """
    if output is None:
        return pptx_path

    # Check if output is a specific .pptx file
    if output.lower().endswith(".pptx"):
        if is_batch and pair_count > 1:
            console.print(
                "[red]Error:[/red] Cannot use -o with a specific .pptx filename "
                "in batch mode with multiple pairs. Use a directory instead."
            )
            sys.exit(1)
        out_path = os.path.abspath(output)
        os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
        shutil.copy2(pptx_path, out_path)
        return out_path

    # Otherwise treat as directory
    out_dir = os.path.abspath(output)
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, os.path.basename(pptx_path))
    shutil.copy2(pptx_path, out_path)
    return out_path


def _make_summary_table(results: dict, column_label: str = "Count") -> Table:
    """Build a rich Table for step results."""
    table = Table(show_header=True)
    table.add_column("Step")
    table.add_column(column_label, justify="right")
    table.add_row("Links", str(results["links"]))
    table.add_row("Charts", str(results["charts"]))
    table.add_row("Tables", str(results["tables"]))
    table.add_row("Deltas", str(results["deltas"]))
    table.add_row("Color coding", str(results["colors"]))
    return table


class _ErrorCollector(logging.Handler):
    """Logging handler that buffers WARNING+ messages for later display."""

    def __init__(self):
        super().__init__(level=logging.WARNING)
        self.errors: list[str] = []

    def emit(self, record):
        self.errors.append(self.format(record))


def process_presentation(
    pptx_path: str,
    excel_path: str,
    config: dict,
    options: argparse.Namespace,
) -> tuple[dict, list[str]]:
    """Process a single presentation through the full pipeline.

    Returns (results_dict, error_messages).
    """
    results = {"links": 0, "tables": 0, "deltas": 0, "colors": 0, "charts": 0}

    # Collect warnings/errors during processing
    collector = _ErrorCollector()
    collector.setFormatter(logging.Formatter("%(message)s"))
    decx_logger = logging.getLogger("decx")
    decx_logger.addHandler(collector)
    old_level = decx_logger.level
    old_propagate = decx_logger.propagate
    verbose = getattr(options, "verbose", False)
    if verbose:
        decx_logger.setLevel(logging.DEBUG)
    else:
        decx_logger.setLevel(logging.WARNING)
        decx_logger.propagate = False

    try:
        with Session(pptx_path, excel_path) as session:
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
    finally:
        decx_logger.removeHandler(collector)
        decx_logger.propagate = old_propagate
        decx_logger.setLevel(old_level)

    return results, collector.errors


def _run_pairs(pairs: list[tuple[str, str]], config: dict, args: argparse.Namespace):
    """Run the pipeline for a list of (pptx_path, excel_path) pairs."""
    grand_total = {"links": 0, "tables": 0, "deltas": 0, "colors": 0, "charts": 0}
    t_start = time.perf_counter()
    processed = 0
    total_files = len(pairs)
    output = getattr(args, "output", None)

    for idx, (pptx_path, excel_path) in enumerate(pairs, 1):
        if not os.path.exists(pptx_path):
            console.print(f"[yellow]PPT not found, skipping:[/yellow] {pptx_path}")
            continue
        if not os.path.exists(excel_path):
            console.print(f"[yellow]Excel not found, skipping:[/yellow] {excel_path}")
            continue

        # Resolve output path (may copy file)
        actual_path = resolve_output_path(
            pptx_path, output, is_batch=len(pairs) > 1, pair_count=len(pairs)
        )

        pptx_name = os.path.basename(pptx_path)
        excel_name = os.path.basename(excel_path)
        verbose = getattr(args, "verbose", False)

        t_file = time.perf_counter()

        if verbose:
            # Verbose: no spinner, just let logs print cleanly
            console.print(
                f"\n[bold]Processing ({idx}/{total_files}):[/bold] "
                f"{pptx_name} <- {excel_name}"
            )
            results, errors = process_presentation(
                actual_path, excel_path, config, args
            )
        else:
            # Normal: spinner with transient (disappears when done)
            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                TimeElapsedColumn(),
                console=console,
                transient=True,
            ) as progress:
                progress.add_task(
                    f"Processing ({idx}/{total_files}): {pptx_name} <- {excel_name}",
                    total=None,
                )
                results, errors = process_presentation(
                    actual_path, excel_path, config, args
                )

        elapsed = time.perf_counter() - t_file

        # Per-file summary
        console.print(f"\n{pptx_name} <- {excel_name} ({elapsed:.2f}s)")
        for err in errors:
            console.print(f"  [bold red]WARNING:[/bold red] {err}")
        console.print(_make_summary_table(results))

        for key in grand_total:
            grand_total[key] += results[key]
        processed += 1

    # Grand total
    total_elapsed = time.perf_counter() - t_start
    console.print(f"\nTotal Summary | {processed} file(s) in {total_elapsed:.2f}s")
    console.print(_make_summary_table(grand_total, column_label="Total"))


def cmd_update(args: argparse.Namespace):
    """Handle the 'update' subcommand — main pipeline."""
    # Logging — suppress ALL console logging by default so spinner stays clean.
    # Errors are captured by _ErrorCollector and shown in red after the spinner.
    # Use -v for full logging output to stderr.
    if args.verbose:
        logging.basicConfig(
            level=logging.DEBUG,
            format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
            datefmt="%H:%M:%S",
        )
    else:
        # Silence all stderr logging — errors go through our collector instead
        logging.basicConfig(level=logging.CRITICAL)

    # Config — defaults + optional --set overrides
    try:
        config = get_config(getattr(args, "set", None))
    except ValueError as e:
        console.print(f"[red]Config error:[/red] {e}")
        sys.exit(1)

    # --- Mode 1: --pair for explicit pptx:xlsx pairs ---
    if args.pair:
        pairs = [parse_pair(p) for p in args.pair]
        _run_pairs(pairs, config, args)
        return

    # --- Mode 2: presentations + --excel (or file picker) ---
    if not args.presentations:
        print("Error: Provide presentation file(s) or use --pair for batch pairs.")
        sys.exit(1)

    excel_path = args.excel
    if not excel_path:
        from decx.file_picker import pick_excel_file

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


def cmd_info(args: argparse.Namespace):
    """Handle the 'info' subcommand — inspect a PPTX file."""
    pptx_path = os.path.abspath(args.presentation)
    if not os.path.exists(pptx_path):
        console.print(f"[red]File not found:[/red] {pptx_path}")
        sys.exit(1)

    with Session(pptx_path, excel_path=None, read_only=True) as session:
        pres = session.presentation
        slide_count = pres.Slides.Count

        # Build inventory
        inventory = build_presentation_inventory(pres)

        # Count unlinked charts by scanning all shapes
        unlinked_charts = _count_all_unlinked_charts(pres)

        # Collect OLE source file paths
        ole_sources: Counter = Counter()
        for _slide, shp in inventory.ole_shapes:
            try:
                source_full = shp.LinkFormat.SourceFullName
                # Split on first '!' to get file path
                file_path = source_full.split("!")[0]
                ole_sources[file_path] += 1
            except Exception:
                ole_sources["(unknown)"] += 1

        # Find template shapes on slide 1
        config = DEFAULT_CONFIG
        template_names = [
            config["delta"]["template_positive"],
            config["delta"]["template_negative"],
            config["delta"]["template_none"],
        ]
        template_found = {}
        for name in template_names:
            shp = find_template_shape(pres, name, slide_index=1)
            template_found[name] = shp is not None

        # Use raw counts (all shapes with each prefix, not just OLE-matched)
        ntbl_count = inventory.count_ntbl
        htmp_count = inventory.count_htmp
        trns_count = inventory.count_trns
        delt_count = inventory.count_delt
        ccst_count = inventory.count_ccst

    # --- Print results ---
    # Presentation table
    console.print("\nPresentation")
    t = Table(show_header=False)
    t.add_column("Key")
    t.add_column("Value")
    t.add_row("File", os.path.basename(pptx_path))
    t.add_row("Slides", str(slide_count))
    console.print(t)

    # OLE Links table
    console.print("\nOLE Links")
    t = Table(show_header=True)
    t.add_column("Source File")
    t.add_column("Count", justify="right")
    total_ole = 0
    for src, count in ole_sources.most_common():
        t.add_row(src, str(count))
        total_ole += count
    t.add_row("Total", str(total_ole), style="bold")
    console.print(t)

    # Charts table
    linked_count = len(inventory.charts)
    console.print("\nCharts")
    t = Table(show_header=True)
    t.add_column("Type")
    t.add_column("Count", justify="right")
    t.add_row("Linked", str(linked_count))
    t.add_row("Unlinked", str(unlinked_charts))
    console.print(t)

    # Special Shapes table
    console.print("\nSpecial Shapes")
    t = Table(show_header=True)
    t.add_column("Type")
    t.add_column("Count", justify="right")
    t.add_row("ntbl_ (normal tables)", str(ntbl_count))
    t.add_row("htmp_ (heatmap tables)", str(htmp_count))
    t.add_row("trns_ (transposed tables)", str(trns_count))
    t.add_row("delt_ (delta indicators)", str(delt_count))
    t.add_row("_ccst (color-coded)", str(ccst_count))
    console.print(t)

    # Delta Templates table
    console.print("\nDelta Templates (Slide 1)")
    t = Table(show_header=True)
    t.add_column("Shape Name")
    t.add_column("Found", justify="center")
    for name in template_names:
        found = "\u2713" if template_found[name] else "\u2717"
        t.add_row(name, found)
    console.print(t)


def _count_unlinked_charts_recursive(shape, results=None):
    """Recursively count charts that are NOT linked."""
    if results is None:
        results = [0]
    from decx.shape_finder import MSO_GROUP

    if shape.Type == MSO_GROUP:
        for sub_shp in shape.GroupItems:
            _count_unlinked_charts_recursive(sub_shp, results)
    elif shape.HasChart:
        try:
            if not shape.Chart.ChartData.IsLinked:
                results[0] += 1
        except Exception:
            pass
    return results[0]


def _count_all_unlinked_charts(presentation) -> int:
    """Count all unlinked charts across the presentation."""
    count = 0
    for slide in presentation.Slides:
        for shp in slide.Shapes:
            count += _count_unlinked_in_shape(shp)
    return count


def _count_unlinked_in_shape(shape) -> int:
    """Count unlinked charts in a shape (recursive for groups)."""
    from decx.shape_finder import MSO_GROUP

    if shape.Type == MSO_GROUP:
        total = 0
        for sub_shp in shape.GroupItems:
            total += _count_unlinked_in_shape(sub_shp)
        return total
    if shape.HasChart:
        try:
            if not shape.Chart.ChartData.IsLinked:
                return 1
        except Exception:
            pass
    return 0


def cmd_config():
    """Handle the 'config' subcommand — show all available --set keys."""
    t = Table(show_header=True)
    t.add_column("Key")
    t.add_column("Default")

    for section, values in DEFAULT_CONFIG.items():
        for key, default in values.items():
            t.add_row(f"{section}.{key}", str(default))

    console.print("\nAvailable --set keys")
    console.print(t)
    console.print("\nUsage: decx update report.pptx -e data.xlsx --set KEY=VALUE")


def main():
    parser = argparse.ArgumentParser(
        prog="decx",
        description="Automated PowerPoint report generation from Excel data via COM",
    )
    parser.add_argument(
        "--version", action="version", version=f"%(prog)s {__version__}"
    )

    subparsers = parser.add_subparsers(dest="command")

    # --- update subcommand ---
    update_parser = subparsers.add_parser(
        "update",
        help="Run the main update pipeline on presentations",
        epilog=(
            "Examples:\n"
            "  decx update report.pptx --excel data.xlsx\n"
            "  decx update report.pptx --excel data.xlsx -o output/\n"
            "  decx update report.pptx --excel data.xlsx -o result.pptx\n"
            "  decx update report.pptx                        (file picker opens)\n"
            '  decx update --pair "us.pptx:us.xlsx" --pair "mx.pptx:mx.xlsx"\n'
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    update_parser.add_argument(
        "presentations",
        nargs="*",
        help="One or more .pptx file paths (supports glob patterns). Used with --excel.",
    )
    update_parser.add_argument(
        "--excel",
        "-e",
        default=None,
        help="Path to the Excel data file. If omitted, a file dialog will open.",
    )
    update_parser.add_argument(
        "--pair",
        "-p",
        action="append",
        default=None,
        metavar="PPT:XLSX",
        help="A pptx:xlsx pair. Can be repeated for batch processing multiple pairs.",
    )
    update_parser.add_argument(
        "--output",
        "-o",
        default=None,
        help=(
            "Output path. If ends with .pptx, write to that file (single-file only). "
            "If a directory, write output files there. If omitted, modify in-place."
        ),
    )
    update_parser.add_argument(
        "--skip-links", action="store_true", help="Skip Step 1a (re-link OLE)"
    )
    update_parser.add_argument(
        "--skip-deltas", action="store_true", help="Skip Step 1c (delta arrows)"
    )
    update_parser.add_argument(
        "--skip-coloring", action="store_true", help="Skip Step 1d (_ccst coloring)"
    )
    update_parser.add_argument(
        "--skip-charts", action="store_true", help="Skip Step 2 (chart links)"
    )
    update_parser.add_argument(
        "--verbose", "-v", action="store_true", help="Enable debug logging"
    )
    update_parser.add_argument(
        "--set",
        action="append",
        default=None,
        metavar="KEY=VALUE",
        help=(
            "Override config value using dot notation. Repeatable. "
            'E.g. --set ccst.positive_prefix="" --set links.set_manual=false'
        ),
    )

    # --- info subcommand ---
    info_parser = subparsers.add_parser(
        "info", help="Inspect a PPTX file and show shape/link inventory"
    )
    info_parser.add_argument(
        "presentation",
        help="Path to the .pptx file to inspect",
    )

    # --- config subcommand ---
    subparsers.add_parser(
        "config", help="Show all available --set keys and their defaults"
    )

    args = parser.parse_args()

    if args.command == "update":
        cmd_update(args)
    elif args.command == "info":
        cmd_info(args)
    elif args.command == "config":
        cmd_config()
    else:
        parser.print_help()
        sys.exit(0)


if __name__ == "__main__":
    main()
