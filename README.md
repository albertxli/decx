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

# Batch mode with explicit pptx:xlsx pairs
decx update --pair "us.pptx:us_data.xlsx" --pair "mx.pptx:mx_data.xlsx"

# Skip specific steps
decx update report.pptx --excel data.xlsx --skip-links --skip-charts

# Use a custom config file
decx update report.pptx --excel data.xlsx --config my_config.yaml

# Verbose output for debugging
decx update report.pptx --excel data.xlsx --verbose
```

### Initialize config

Write the default `config.yaml` to the current directory:

```bash
decx init
```

### Info

```bash
decx info
```

### Version

```bash
decx --version
```

## Configuration

`decx` ships with sensible defaults. Run `decx init` to generate a `config.yaml` you can customize:

```yaml
heatmap:
  color_minimum: '#F8696B'
  color_midpoint: '#FFEB84'
  color_maximum: '#63BE7B'
  dark_font: '#000000'
  light_font: '#FFFFFF'

ccst:
  positive_color: '#33CC33'
  negative_color: '#ED0590'
  neutral_color: '#595959'
  positive_prefix: '+'
  symbol_removal: '%'

delta:
  template_positive: tmpl_delta_pos
  template_negative: tmpl_delta_neg
  template_none: tmpl_delta_none
  template_slide: 1

links:
  set_manual: true
```

## Pipeline

1. **Re-link OLE objects** — point linked Excel objects to a new data file
2. **Populate tables** — read Excel ranges and write values into PowerPoint tables
3. **Delta indicators** — swap arrow shapes based on positive/negative values
4. **Color coding** — apply color rules to `_ccst` tables
5. **Update charts** — refresh linked chart data sources

## License

MIT

## Repository

https://github.com/albertxli/decx
