# decx

Automated PowerPoint report generation from Excel data via COM.

`decx` reads data from Excel workbooks and updates linked OLE objects, tables, delta indicators, color coding, and charts in PowerPoint presentations — all driven from the command line.

## Requirements

- **Windows** (COM automation requires Windows)
- **Microsoft PowerPoint** (installed and licensed)
- **Microsoft Excel** (installed and licensed)
- **Python 3.11+**

## Installation

```bash
uv add decx
```

Or with pip:

```bash
pip install decx
```

## Usage

### Update presentations

```bash
# Single presentation with one Excel file
decx update report.pptx --excel data.xlsx

# Save output to a specific file (instead of modifying in-place)
decx update report.pptx --excel data.xlsx -o updated_report.pptx

# Save output to a directory
decx update report.pptx --excel data.xlsx -o output/

# Batch mode with explicit pptx:xlsx pairs
decx update --pair "us.pptx:us_data.xlsx" --pair "mx.pptx:mx_data.xlsx"

# Batch mode with output directory
decx update --pair "us.pptx:us.xlsx" --pair "mx.pptx:mx.xlsx" -o output/

# Skip specific steps
decx update report.pptx --excel data.xlsx --skip-links --skip-charts

# Override config values
decx update report.pptx --excel data.xlsx --set ccst.positive_prefix=""

# Verbose output for debugging
decx update report.pptx --excel data.xlsx --verbose
```

The `-o`/`--output` flag controls where results are written:
- **`-o result.pptx`** — write to a specific file (single-file mode only)
- **`-o output/`** — write to a directory (works with batch mode too)
- **omitted** — modify the source file in-place

After processing, `decx update` prints a summary table per file and a grand total:

```
report.pptx <- data.xlsx (19.38s)
+--------------+-------+
| Step         | Count |
+--------------+-------+
| Links        |    86 |
| Tables       |    86 |
| Deltas       |     0 |
| Color coding |     0 |
| Charts       |   100 |
+--------------+-------+

All done! 1 file(s) in 19.38s
+--------------+-------+
| Step         | Total |
+--------------+-------+
| Links        |    86 |
| Tables       |    86 |
| Deltas       |     0 |
| Color coding |     0 |
| Charts       |   100 |
+--------------+-------+
```

### Inspect a presentation

```bash
decx info report.pptx
```

Sample output:

```
Presentation
+----------+----------------------------------------+
| File     | report.pptx                            |
| Slides   | 30                                     |
+----------+----------------------------------------+

OLE Links
+-------------------------------------------+-------+
| Source File                               | Count |
+-------------------------------------------+-------+
| C:\data\tracking_Argentina.xlsx           |    86 |
| Total                                     |    86 |
+-------------------------------------------+-------+

Charts
+-------------------------------------------+-------+
| Type                                      | Count |
+-------------------------------------------+-------+
| Linked                                    |   100 |
| Unlinked                                  |     0 |
+-------------------------------------------+-------+

Special Shapes
+-------------------------------------------+-------+
| Type                                      | Count |
+-------------------------------------------+-------+
| ntbl_ (normal tables)                     |    42 |
| htmp_ (heatmap tables)                    |     3 |
| trns_ (transposed tables)                 |     2 |
| delt_ (delta indicators)                  |     8 |
| _ccst (color-coded)                       |     5 |
+-------------------------------------------+-------+

Delta Templates (Slide 1)
+-------------------------------------------+-------+
| Shape Name                                | Found |
+-------------------------------------------+-------+
| tmpl_delta_pos                            |   v   |
| tmpl_delta_neg                            |   v   |
| tmpl_delta_none                           |   v   |
+-------------------------------------------+-------+
```

### Version

```bash
decx --version
```

## CLI Reference

```
decx --version                  Show version and exit
decx --help                     Show help

decx update [FILES] [OPTIONS]   Run the update pipeline
  FILES                         One or more .pptx files (glob patterns OK)
  -e, --excel PATH              Excel data file (or file picker opens)
  -p, --pair PPT:XLSX           Explicit pptx:xlsx pair (repeatable)
  -o, --output PATH             Output file (.pptx) or directory
  --skip-links                  Skip OLE re-linking
  --skip-deltas                 Skip delta indicator updates
  --skip-coloring               Skip _ccst color coding
  --skip-charts                 Skip chart link updates
  -v, --verbose                 Debug logging
  --set KEY=VALUE               Override config value (repeatable, dot notation)

decx info FILE                  Inspect a .pptx file (read-only, no Excel needed)

decx config                     Show all available --set keys and defaults
```

## Configuration

`decx` ships with sensible defaults. Use `--set` to override values, or run `decx config` to see all available keys.

Any config value can also be overridden from the CLI with `--set`:

```bash
decx update report.pptx -e data.xlsx --set ccst.positive_prefix="" --set links.set_manual=false
```

### Available `--set` keys

| Key | Default | Description |
|---|---|---|
| `heatmap.color_minimum` | `#F8696B` | Heatmap low color (red) |
| `heatmap.color_midpoint` | `#FFEB84` | Heatmap mid color (yellow) |
| `heatmap.color_maximum` | `#63BE7B` | Heatmap high color (green) |
| `heatmap.dark_font` | `#000000` | Dark font for light heatmap cells |
| `heatmap.light_font` | `#FFFFFF` | Light font for dark heatmap cells |
| `ccst.positive_color` | `#33CC33` | Font color for positive numbers |
| `ccst.negative_color` | `#ED0590` | Font color for negative numbers |
| `ccst.neutral_color` | `#595959` | Font color for zero / non-numeric |
| `ccst.positive_prefix` | `+` | Prefix for positive numbers (set `""` to disable) |
| `ccst.symbol_removal` | `%` | Characters to strip after coloring (any combo of `+`, `-`, `%`) |
| `delta.template_positive` | `tmpl_delta_pos` | Positive delta template shape name |
| `delta.template_negative` | `tmpl_delta_neg` | Negative delta template shape name |
| `delta.template_none` | `tmpl_delta_none` | No-change delta template shape name |
| `delta.template_slide` | `1` | Slide number where delta templates live |
| `links.set_manual` | `true` | Set OLE links to manual update mode |

## Pipeline

1. **Re-link OLE objects** — point linked Excel objects to a new data file
2. **Populate tables** — read Excel ranges and write values into PowerPoint tables
3. **Delta indicators** — swap arrow shapes based on positive/negative values
4. **Color coding** — apply color rules to `_ccst` tables
5. **Update charts** — refresh linked chart data sources

## Benchmark

Reference benchmarks on a 30-slide presentation with 86 OLE objects and 100 charts. Actual speed varies by machine, file size, and disk speed.

| Scenario | Time |
|---|---|
| Same Excel file (re-run) | ~17s |
| Different Excel file (same folder) | ~36s |
| Different Excel file (remote/network folder) | ~36s |

Batch processing 3 country reports: ~60s total.

> **Note:** OLE visual thumbnails are not refreshed during update (they show cached images). This is intentional — refreshing each OLE visual adds ~4s per object for remote files and provides no value since OLE objects are hidden behind native PowerPoint tables. Tables, charts, and delta indicators all receive fresh data directly.

## License

MIT

## Repository

https://github.com/albertxli/decx
