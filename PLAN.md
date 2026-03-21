# Implementation Plan Log

## Phase 1: Base Functionality (Complete)

### Goal
Port the VBA macro (`RUN ALL_Table+Chart_v11.bas`, ~1700 lines) to Python using `win32com`.

### Steps Completed
1. **Foundation modules** — `utils.py` (hex→RGB, R1C1→A1, link parsing, contrast color), `shape_finder.py` (token matching, shape discovery)
2. **COM session manager** — `session.py` with `DispatchEx` for isolated processes, auto-dismiss security dialog via `win32gui` background thread
3. **Config** — `config.yaml` with all user-configurable settings (colors, prefixes, template names)
4. **Step 1a** — `linker.py` re-points OLE links to new Excel file, sets links to manual
5. **Formatting** — `formatting.py` extracts/applies table formatting (fonts, borders, fills, margins)
6. **Step 1b** — `table_updater.py` populates PPT tables from Excel data (ntbl_, htmp_, trns_)
7. **Step 1c** — `delta_updater.py` swaps delta indicator arrows (two-pass approach)
8. **Step 1d** — `color_coder.py` applies _ccst sign-based color coding
9. **Step 2** — `chart_updater.py` updates chart links
10. **CLI** — `main.py` with argparse, batch support (`--pair`), file picker fallback

### Key Issues Resolved
- `msoLinkedOLEObject = 10` not 7 (GOTCHAS #2)
- "Update Links" dialog requires `win32gui` auto-dismiss (GOTCHAS #1)
- `ppUpdateOptionManual = 1` not 2 (GOTCHAS #11)
- COM zombie cleanup with `gc.collect()` + `time.sleep(1)` (GOTCHAS #3)

### Test Results
- **58 unit tests** — all passing (0.08s)
- **3 integration tests** — all passing (~197s total)
- **Runtime**: ~60s per presentation (30 slides, 86 OLE objects)

---

## Phase 2: Performance Optimization (In Progress)

### Goal
Reduce pipeline runtime from ~60s to ~15-30s by cutting ~200k+ redundant COM calls.

### Problem Analysis
Current COM call breakdown (~307k total):
| Operation | COM Calls | % of Total |
|---|---|---|
| Formatting extract/apply | ~232,200 | 75% |
| Cell text/color copy | ~60,200 | 20% |
| Shape finding (3-pass) | ~7,740 | 3% |
| Color coding (_ccst) | ~4,400 | 1% |
| Link/chart updates | ~3,000 | 1% |

### Optimization Steps

#### Step 1: Single-Pass Shape Inventory
**Files:** `shape_finder.py`, all step modules, `main.py`
**What:** Build one index of all shapes per slide (OLE, tables, delts, ccst, charts) in a single pass. All modules query the index instead of scanning shapes independently.
**Saves:** ~8,000 COM calls, ~2-4s

#### Step 2: Defer LinkFormat.Update() to Bulk
**Files:** `linker.py`
**What:** Set all `SourceFullName` paths first, then call `presentation.UpdateLinks()` once instead of 86 individual `LinkFormat.Update()` calls. OLE visuals still get refreshed.
**Saves:** 85 slow refresh operations, ~10-20s

#### Step 3: Smarter Formatting — Text-Only for Preserved Tables
**Files:** `table_updater.py`, `formatting.py`
**What:** For `ntbl_` and `trns_` tables (where `PreserveFormattedColor = True`), only update cell text — skip the full 27-property extract/apply cycle. Formatting is already correct from the previous run. Only `htmp_` tables (which re-pull colors from Excel) need full formatting.
**Saves:** ~150,000-230,000 COM calls, ~10-15s

#### Step 4: Remove time.sleep(1)
**Files:** `session.py`
**What:** Remove the 1-second sleep in `__exit__`. Test if `gc.collect()` alone is sufficient.
**Saves:** 1s per file

#### Step 5: Batch Excel Reads with Range.Value2
**Files:** `table_updater.py`
**What:** Read entire Excel range in one COM call (`Range.Value2` returns 2D tuple) instead of per-cell `Cells(row, col).Text`.
**Saves:** ~17,000 COM calls, ~3-5s

### Expected Results
| Metric | Phase 1 | Phase 2 Target |
|---|---|---|
| Runtime (30 slides, 86 OLE) | ~60s | ~15-30s |
| COM calls | ~307,000 | ~50,000-100,000 |
| Improvement | baseline | 50-75% faster |

### Verification
After each optimization step:
1. `uv run pytest tests/ -k "not integration"` — unit tests pass
2. `uv run pytest tests/ -k integration` — integration tests pass
3. Open output .pptx and verify data/formatting matches Phase 1 output
4. Runtime benchmark: time the full pipeline on test template + Argentina Excel

### Benchmark Tests (to add)
- `tests/test_benchmark.py` — timed integration tests comparing Phase 1 vs Phase 2 runtime
  - `test_benchmark_single_presentation` — time full pipeline on template + one Excel file
  - `test_benchmark_batch_three_countries` — time 3-pair batch run
  - Print elapsed times so we can track improvements
