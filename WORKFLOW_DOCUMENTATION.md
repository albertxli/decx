# PowerPoint Excel Data Automation — Workflow Documentation

## Overview

This VBA macro automates the process of updating PowerPoint presentations with data from Excel workbooks. It is designed for multi-market reporting where the same presentation template is reused across markets by swapping the underlying Excel data source.

---

## Execution Pipeline

### `Step_1_RUN_ALL` (Master Wrapper)

Runs all steps in sequence with silent mode enabled (no per-step popups). Displays a single summary at the end.

```
Step_1a  Re-link OLE objects to a new Excel file
   |
Step_1b  Refresh/rebuild PowerPoint-native tables from Excel data
   |        (keeps Excel instance alive in batch mode)
   |
Step_1c  Apply delta indicator arrows based on linked values
   |        (reuses shared Excel instance, then cleans up)
   |
Step_1d  Apply +/- color coding to tables with _ccst suffix
   |        (no Excel needed — pure PowerPoint)
   |
Step_2   Update linked chart data sources
```

Each step can also be run **standalone** (outside of RUN_ALL).

---

## Step Details

### Step_1a — `Step_1a_UpdateTableLinks`

**Purpose**: Re-point all linked OLE objects from the old Excel file to a new one.

- Prompts the user to select a new Excel file
- Iterates all slides and collects linked OLE objects (including inside groups)
- Changes each `LinkFormat.SourceFullName` to the new file path while preserving the sheet and range
- Calls `LinkFormat.Update` per shape to refresh the OLE visual
- Sets links to manual update mode (`ppUpdateOptionManual`)
- Stores the selected file path in `gSelectedExcelFile` for reuse by Step_2

**Toggles** (inside the sub):
| Variable | Default | Description |
|----------|---------|-------------|
| `doSinglePassUpdate` | `True` | *(Currently disabled)* Bulk-updates all links at the end |
| `setLinksToManual` | `True` | Sets each link's AutoUpdate to manual |

---

### Step_1b — `Step_1b_UpdateTableContent`

**Purpose**: Read data from Excel workbooks and populate PowerPoint-native tables.

For each linked OLE object on every slide:
1. Extracts file path, sheet name, and cell range from the OLE's `SourceFullName`
2. Looks for an existing special-type shape (`ntbl_`, `htmp_`, or `trns_`) associated with this OLE
3. If a `delt_`-only shape exists (no table type) — skips this OLE (no table needed)
4. Opens the Excel workbook and reads the cell range
5. Creates or updates a PowerPoint-native table with the Excel data
6. Preserves formatting (fonts, borders, fills, margins) from the existing table

**Batch mode**: Stores the Excel instance in `gExcelApp` / `gOpenedWorkbooks` globals so Step_1c can reuse it without reopening.

---

### Step_1c — `Step_1c_ApplyDeltaIndicators`

**Purpose**: Swap delta indicator arrow shapes based on linked Excel values (positive/negative/no change).

**How it works**:
1. Locates 3 template shapes on **Slide 1**: `tmpl_delta_pos`, `tmpl_delta_neg`, `tmpl_delta_none`
2. **Pass 1** (collect): Iterates all OLE objects, finds those with matching `delt_` shapes. Saves position, size, and name. No shapes are modified (safe for iteration).
3. **Pass 2** (process): For each collected item:
   - **Primary**: Reads value from the corresponding `ntbl_`/`htmp_`/`trns_` table (if it exists)
   - **Fallback**: Opens Excel via shared `gExcelApp` (batch mode) or lazy-creates a new instance (standalone)
   - Determines sign: positive (>0), negative (<0), no change (=0). Handles `%` suffix.
   - Deletes the old `delt_` shape
   - Copies the correct template arrow from Slide 1
   - Restores position, size, and name

**Template shapes** (user must create on Slide 1):
| Shape Name | Meaning |
|------------|---------|
| `tmpl_delta_pos` | Green up-arrow (positive change) |
| `tmpl_delta_neg` | Brown/dark down-arrow (negative change) |
| `tmpl_delta_none` | Grey bowtie/hourglass (no change) |

---

### Step_1d — `Step_1d_ColorNumbersInTables`

**Purpose**: Apply color coding to numeric cells in tables whose names contain `_ccst`.

**Behavior**:
- Scans all shapes on all slides for names containing `_ccst`
- For each matching table, iterates every cell:
  - **Positive values** (>0): colored green (`#33CC33`), optional `+` prefix added
  - **Negative values** (<0): colored red (`#ED0590`)
  - **Zero or text**: colored grey (`#595959`)
- Optionally strips symbols (`%`, `+`, `-`) after coloring

**Customizable settings** (inside the sub):
| Variable | Default | Description |
|----------|---------|-------------|
| `positiveHex` | `#33CC33` | Font color for positive numbers |
| `negativeHex` | `#ED0590` | Font color for negative numbers |
| `neutralHex` | `#595959` | Font color for zero / non-numeric text |
| `positivePrefix` | `"+"` | Prefix added to positive numbers (set `""` to disable) |
| `symbolRemoval` | `"%"` | Characters to strip after coloring. Any combo of `+`, `-`, `%` |

---

### Step_2 — `Step_2_UpdateChartLinks`

**Purpose**: Re-link embedded charts to the selected Excel file.

- In batch mode, reuses the file path from `gSelectedExcelFile` (set by Step_1a)
- In standalone mode, prompts the user to pick an Excel file
- Iterates all slides, collects linked charts (including inside groups)
- Updates each chart's `LinkFormat.SourceFullName` and calls `LinkFormat.Update`

---

## Special Object Types

PowerPoint shapes are identified by **name prefixes**. The macro uses word-boundary matching to associate OLE objects with their corresponding special shapes.

### `ntbl_` — Normal Table

- **Created by**: Step_1b (auto-created for any OLE without an existing special shape)
- **Naming**: `ntbl_<OLE object name>` (e.g., `ntbl_Object 5`)
- **Behavior**: Copies Excel cell values into a PowerPoint-native table. **Preserves** existing formatting (fonts, borders, fills, margins) on subsequent runs.
- **Use case**: Standard data tables that need consistent formatting across market swaps

### `htmp_` — Heatmap Table

- **Created by**: User (rename an existing `ntbl_` to `htmp_`)
- **Naming**: `htmp_<OLE object name>`
- **Behavior**: Re-applies a **3-color scale** (red/yellow/green) from Excel on every run. Does NOT preserve cell fill colors — recalculates them fresh. Applies contrast font colors (black/white) based on background brightness.
- **Color constants** (global):
  - `COLOR_MINIMUM` = `#F8696B` (red)
  - `COLOR_MIDPOINT` = `#FFEB84` (yellow)
  - `COLOR_MAXIMUM` = `#63BE7B` (green)
- **Use case**: Heatmap visualizations where colors must reflect the current data

### `trns_` — Transposed Table

- **Created by**: User (rename an existing `ntbl_` to `trns_`)
- **Naming**: `trns_<OLE object name>`
- **Behavior**: Same as `ntbl_` (preserves formatting) but **transposes** the data — rows become columns and vice versa when filling cells.
- **Use case**: When the Excel range is row-oriented but the PowerPoint table should be column-oriented (or vice versa)

### `delt_` — Delta Indicator

- **Created by**: User (manually place a template arrow shape and rename it)
- **Naming**: `delt_<OLE object name>`
- **Behavior**: NOT a table. Reads a single-cell value from the linked OLE object. Based on the sign, swaps the shape with the correct template arrow from Slide 1. **Preserves position and size**.
- **Important**: If no `delt_` shape exists for an OLE → nothing happens (unlike `ntbl_` which auto-creates). Step_1b will also skip creating an `ntbl_` for delt_-only OLE objects.
- **Use case**: Year-over-year delta indicators next to charts (positive/negative/no change arrows)

---

## Name-Based Features

### `_ccst` Suffix — Color-Coded Sign Table

- **Trigger**: Any table shape with `_ccst` anywhere in its name
- **Processed by**: Step_1d
- **Effect**: Numeric cells get colored by sign (green/red/grey) with optional `+` prefix and symbol stripping
- **Can combine**: e.g., `ntbl_Object5_ccst` — Step_1b updates the data, Step_1d colors it

---

## Word-Boundary Matching

All shape-finding functions (`FindExistingNtblTable`, `FindExistingHtmpTable`, `FindExistingTrnsTable`, `FindExistingDeltShape`) use **exact token matching** via `IsExactTokenMatch`.

This prevents false positives where a shorter name is a substring of a longer one. A match requires the OLE object name to appear as a **complete token** in the shape name, bounded by:
- Start/end of string
- Underscore `_`
- Space
- Hyphen `-`
- Any non-alphanumeric character

**Examples**:
| Shape Name | OLE Name | Match? |
|------------|----------|--------|
| `ntbl_Object_5` | `Object_5` | Yes |
| `ntbl_Object_55` | `Object_5` | No (5 is followed by another digit) |
| `delt_NetSent_ccst` | `NetSent` | Yes |

---

## Architecture: Shared Excel Instance

In batch mode (`Step_1_RUN_ALL`), the Excel instance is shared across steps to avoid COM zombie errors and improve performance:

```
Step_1b: Creates Excel instance → stores in gExcelApp / gOpenedWorkbooks
         Processes ntbl_/htmp_/trns_ tables
         Skips cleanup (batch mode)
            |
Step_1c: Reuses gExcelApp for delt_-only OLE objects
         Reads from ntbl_ tables when available (no Excel needed)
         Cleans up: closes workbooks, quits Excel, clears globals
            |
Step_1d: No Excel needed (pure PowerPoint)
            |
Step_2:  Creates its own separate Excel instance for charts
```

In **standalone mode** (running a step individually), each step manages its own Excel lifecycle independently.

---

## Setup Checklist

1. **Template arrows**: Place 3 shapes on Slide 1 named `tmpl_delta_pos`, `tmpl_delta_neg`, `tmpl_delta_none`
2. **OLE links**: Paste-link Excel ranges into PowerPoint slides
3. **Name special shapes**: Rename PowerPoint-native tables/shapes with the appropriate prefix:
   - `ntbl_<OLE name>` for normal tables
   - `htmp_<OLE name>` for heatmap tables
   - `trns_<OLE name>` for transposed tables
   - `delt_<OLE name>` for delta indicators
4. **Add `_ccst`** to any table name that should get color-coded numbers
5. **Run**: Execute `Step_1_RUN_ALL` or individual steps as needed
