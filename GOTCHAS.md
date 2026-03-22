# Gotchas & Lessons Learned

Hard-won knowledge from debugging COM automation with PowerPoint and Excel via pywin32. **Read this before making changes.**

---

## 1. "Update Links" Security Dialog Blocks COM

**Problem:** When opening a `.pptx` file that contains OLE links via `Presentations.Open()`, PowerPoint shows a "Microsoft PowerPoint Security Notice" dialog asking whether to update links. This dialog is **modal and blocking** — your Python script hangs until someone clicks a button.

**What does NOT work:**
- `ppt.DisplayAlerts = 0` — does not suppress this specific dialog
- `ppt.AutomationSecurity = 3` (msoAutomationSecurityForceDisable) — does not suppress it
- Registry key `WorkbookLinkWarnings` — does not suppress it
- Any combination of the above

**What works:** A background thread using `win32gui.EnumWindows()` that polls for the dialog by window title and sends `WM_CLOSE` to dismiss it. Start the thread **before** calling `Presentations.Open()`.

See: `decx/session.py` — `_auto_dismiss_security_dialog()`

---

## 2. MSO_LINKED_OLE_OBJECT = 10, Not 7

**Problem:** The `msoLinkedOLEObject` constant is `10`, not `7`. Using `7` (which is `msoPlaceholder`) causes all OLE shape detection to silently fail — no errors, just zero shapes found.

**Correct constants:**
- `msoLinkedOLEObject = 10`
- `msoGroup = 6`

Always verify COM constants against Microsoft docs or by inspecting in PowerPoint's VBA editor (`?msoLinkedOLEObject`).

---

## 3. COM Zombie Errors During Cleanup (0x80010108)

**Problem:** After calling `ppt.Quit()`, subsequent COM calls on the dead object (like `presentation.Close()` or `ppt.Presentations.Count`) trigger `Windows fatal exception: code 0x80010108` (RPC_E_DISCONNECTED). These are C-level crashes that Python's `try/except` catches but Windows still prints to stderr.

**Solution:**
- Quit Excel **before** PowerPoint (Excel cleanup is less noisy)
- Call `ppt.Quit()` **last**, don't try to call anything on the object after
- Use `gc.collect()` + `time.sleep(1)` after cleanup to let OS release processes
- The error messages are **cosmetic** — tests still pass. Don't chase them.

---

## 4. Dispatch vs DispatchEx vs EnsureDispatch

**Problem:** Different dispatch methods have different behaviors:

| Method | Behavior |
|---|---|
| `Dispatch("PowerPoint.Application")` | Connects to existing instance OR creates new. **Dangerous** — may interfere with user's open files |
| `DispatchEx("PowerPoint.Application")` | Always creates a **new isolated** process. Safe for automation. |
| `gencache.EnsureDispatch(...)` | Like Dispatch but generates Python type wrappers for better IntelliSense. Still connects to existing instance. |

**Rule:** Use `DispatchEx` for automation scripts to avoid touching user's open applications.

---

## 5. Excel UpdateLinks Parameter

**Problem:** When opening an Excel workbook via COM, Excel may show its own "Update Links" dialog or auto-refresh external data, which slows things down.

**Solution:** Always pass `UpdateLinks=0` when opening workbooks:
```python
wb = excel_app.Workbooks.Open(file_path, UpdateLinks=0, ReadOnly=False)
```

---

## 6. Excel Calculation Mode for Performance

**Problem:** Excel recalculates formulas on every cell change during COM automation, which is extremely slow.

**Solution:** Set calculation to manual before processing, restore after:
```python
excel_app.Calculation = -4135  # xlCalculationManual
# ... do work ...
excel_app.Calculation = -4105  # xlCalculationAutomatic
```

---

## 7. Floating Point in Brightness Calculation

**Problem:** The contrast font color formula `0.299*128 + 0.587*128 + 0.114*128` should equal `128.0` but Python returns `127.99999999999999` due to IEEE 754 floating point. This affects the `< 128` threshold comparison.

**Impact:** Matches VBA behavior (same floating point), so it's not a bug — but don't write tests that assume exact `128.0`.

---

## 8. VBA File is the Authoritative Reference

**Problem:** There is an old Jupyter notebook (`PPT chart update python.ipynb`) with a 2024 Python attempt. Its functionality is **outdated** and does not match the current VBA code.

**Rule:**
- `RUN ALL_Table+Chart_v11.bas` = authoritative source for all logic and functionality
- Old Jupyter notebook = only useful for pywin32 COM patterns (dispatch methods, constant values)
- Never change functionality based on the old Python code

---

## 9. Zombie Processes After Failed Test Runs

**Problem:** If a test crashes or is killed, PowerPoint/Excel processes remain as zombies in Task Manager. The next test run may fail with `RPC_E_CALL_REJECTED (0x80010001)` because it tries to connect to the zombie.

**Solution:** Before running tests after a crash:
```bash
taskkill /F /IM POWERPNT.EXE
taskkill /F /IM EXCEL.EXE
```

---

## 10. pyproject.toml Corruption with uv

**Problem:** Running `uv add` while `pyproject.toml` has a syntax error (e.g., unclosed array) can truncate the file further.

**Solution:** Always verify `pyproject.toml` is valid TOML before running `uv add`. If corrupted, rewrite the entire file.

---

## 11. ppUpdateOptionManual = 1, NOT 2

**Problem:** The PowerPoint constant `ppUpdateOptionManual` is `1`, and `ppUpdateOptionAutomatic` is `2`. Setting `AutoUpdate = 2` (thinking it's manual) actually sets links to **automatic**, which causes the "Update Links" dialog to appear every time the user opens the file.

**Verified via:**
```python
from win32com.client import constants
print(constants.ppUpdateOptionManual)    # 1
print(constants.ppUpdateOptionAutomatic) # 2
```

**Rule:** Always verify COM constants from the type library, not from documentation or guesses. Use `win32com.client.constants` after `gencache.EnsureDispatch()`.

**Impact:** ALL links (OLE worksheet objects AND charts) must be set to `AutoUpdate = 1` before saving, otherwise the security dialog appears on every open.

---

## 12. Bulk UpdateLinks() Can Be SLOWER Than Per-Shape Updates

**Problem:** `presentation.UpdateLinks()` refreshes ALL links in one call, which sounds faster. But in practice with 86 OLE objects, it took ~230s per test vs ~50s with per-shape `LinkFormat.Update()`.

**Why:** The bulk call appears to do a full presentation-wide refresh including re-rendering all OLE visuals, while per-shape updates are more targeted.

**Rule:** Use per-shape `shp.LinkFormat.Update()` after repointing each link, not bulk `presentation.UpdateLinks()`.

---

## 13. delt_ Shapes Can Be Groups (msoGroup = 6)

**Problem:** Delta indicator arrows (`delt_`) are often group shapes (type=6) containing multiple sub-shapes (arrow + styling). The shape scanning code recursed into groups but returned early, so the group's own name was never checked for `delt_`. This caused `inventory.delts` to miss grouped delt_ shapes, and table_updater would auto-create unwanted `ntbl_` tables for delt_-only OLE objects.

**Fix:** In `_scan_shape_recursive()`, check the group shape's name for `delt_` BEFORE recursing into its children.

**Rule:** Any shape type prefix check (`delt_`, `ntbl_`, etc.) must happen on the shape itself, not just its children. Groups are first-class shapes with names.

---

## 14. Rich Library Unicode in Claude Code's Terminal vs User's Bash

**Problem:** Claude Code's terminal uses cp1252 encoding, which can't render rich's Unicode spinner characters. This caused `UnicodeEncodeError` during development.

**Reality:** The user runs bash which handles Unicode fine. The cp1252 issue is only in Claude Code's testing environment, not the user's actual terminal.

**Rule:** Use Unicode freely in the CLI (spinners, checkmarks, etc.). If Claude Code's terminal crashes on Unicode output, that's a testing limitation — the actual user won't see it.

---

## 15. LinkFormat.Update() is Extremely Slow for Remote Excel Files

**Problem:** `shp.LinkFormat.Update()` forces PowerPoint to fetch data from the Excel file and re-render the OLE visual. For each of 86 shapes, this takes ~0.1s for local files but **~4s for files in a different folder** (especially Synology Drive/network paths). Total: 86 x 4s = **344s (5.7 minutes)** just for the link update step.

**Why:** PowerPoint launches its own Excel connection per update. When the file path is remote or on a slow drive, each connection has latency.

**Fix:** Skip `LinkFormat.Update()` entirely. Just repoint `SourceFullName` and set `AutoUpdate = 1` (manual). OLE visuals show stale cached images, but they're hidden behind ntbl_ tables anyway. Our pipeline populates tables and charts with fresh data.

**Impact:** Remote file: 350s+ (frozen) -> 35s. Local file: saves ~3-4s.

**Rule:** Never call `LinkFormat.Update()` — it's the single biggest performance killer for remote files.

---

## 16. Range.Value2 Loses Number Formatting — Use .Text Instead

**Problem:** `Range.Value2` returns raw numeric values (0.05 for "5%", 1234.5 for "1,234.5"). Using it to populate PPT tables destroys all Excel number formatting — percentages show as decimals, formatted numbers lose separators. It can also return `None` for empty ranges, causing tables to silently not update.

**What we tried:** Bulk-read with `Range.Value2` (one COM call) + custom `_format_value()` to convert. This saved ~0.5s per file but broke formatting for all tables.

**Fix:** Reverted to per-cell `cell_range.Cells(row, col).Text` which returns the exact formatted display string from Excel. The 0.5s performance cost is worth preserving correct number formatting.

**Rule:** Always use `.Text` for reading cell values destined for display. `.Value2` is only safe for numeric calculations where formatting doesn't matter.

---

## 17. Inventory Must Be Keyed Per-Slide — Same OLE Names on Different Slides

**Problem:** Inventory dicts (`tables`, `delts`) were keyed by OLE name alone (`str`). When multiple slides have OLE objects with the same name (e.g., `Object_23n` on slides 6-27), only the first slide's match was stored. All subsequent slides with the same OLE name silently skipped table updates.

**Why it happens:** The template uses consistent naming — the same OLE name appears on many slides, each linked to a different Excel range. This is by design.

**Fix:** Key inventory dicts by `(slide_index, ole_name)` tuple. Match table/delt shapes to OLE objects only on the same slide. This matches VBA behavior where `FindExistingNtblTable` takes the slide as a parameter.

**Rule:** Never use bare OLE names as dict keys in the inventory. Always include the slide index to handle duplicate names across slides.

---

## 18. Series.Formula Is Inaccessible on Linked Charts

**Problem:** `Chart.SeriesCollection(i).Formula` throws a COM error (`OLE error 0xe0000002`) on linked charts in PowerPoint. This property works fine on embedded/unlinked charts and in Excel, but fails on charts linked to an external workbook.

**What does NOT work:**
- `Series.Formula` — COM error
- `Series.FormulaR1C1` — same error
- `ChartData.Activate()` then reading `Series.Formula` — opens a **visible Excel window** per chart (~4s each), often hangs or crashes the RPC server. Completely unusable for 100+ charts.
- `ChartData.Workbook` — requires `Activate()` first (same problem)

**What DOES work:**
- `Series.Values` — returns tuple of floats (the cached plotted values), no activation needed
- `Series.XValues` — returns tuple of category labels
- `Series.Name` — returns the series name
- Parsing the PPTX as a zip file and reading the chart XML directly

**Workaround:** Open the `.pptx` as a zip archive. Chart data references are stored in `ppt/charts/chartN.xml` inside `<c:numRef><c:f>` elements (e.g. `Tables!$B$9:$B$13`). No COM needed for this step — pure XML parsing.

**Rule:** Never use `ChartData.Activate()` for reading chart data. Parse the PPTX zip for range references, and use `Series.Values` via COM for the actual plotted values.

See: `decx/checker.py` — `_build_chart_ref_map()`, `_read_chart_range()`

---

## 19. Chart-to-XML Mapping Must Use Slide+Position, Not Flat Index

**Problem:** `inventory.charts` (COM) and `ppt/charts/chartN.xml` files are NOT in the same order. Chart XML files are numbered sequentially (`chart1.xml` through `chart100.xml`), but COM chart shapes have arbitrary names like `Chart 7`, `Chart 33`. Matching by flat index produces completely wrong range references — every chart compares against the wrong Excel data.

**Why it happens:** Charts are created, deleted, and re-added during template editing. The XML numbering reflects creation order, not slide order. COM iterates shapes in slide presentation order.

**What works:** Map via the slide XML relationship chain:
1. Parse each `ppt/slides/slideN.xml` — find `<p:graphicFrame>` elements containing `<c:chart r:id="rIdX"/>` in document order
2. Look up `rIdX` in `ppt/slides/_rels/slideN.xml.rels` → get target like `../charts/chart5.xml`
3. Parse that chart XML for `<c:numRef><c:f>` range references
4. Key by `(slide_number, chart_position_on_slide)` — this matches COM iteration order

**Rule:** Never match charts by flat index or XML filename number. Always use the slide relationship chain to build a `(slide, position)` keyed map.

See: `decx/checker.py` — `_build_chart_ref_map()`

---

## 20. Chart Series Can Reference Non-Contiguous Excel Ranges

**Problem:** Some chart series pull data from non-adjacent cells. The range reference in the chart XML looks like `(Tables!$C$810,Tables!$F$810)` — two separate cells joined by a comma, wrapped in parentheses. Excel COM's `Range()` method rejects comma-separated ranges and throws an exception.

**Why it happens:** Charts built from non-adjacent columns (e.g. comparing two different time periods that aren't in consecutive columns) store the multi-area reference as a comma-delimited string.

**What does NOT work:**
- `wb.Sheets("Tables").Range("C810,F810")` — COM exception

**What works:** Split the reference on commas, read each sub-range separately, and concatenate the values:
```python
sub_refs = ref.split(",")  # ["Tables!$C$810", "Tables!$F$810"]
for sub_ref in sub_refs:
    cell_range = wb.Sheets(sheet).Range(addr)
    values.extend(...)
```

Also strip outer parentheses if present: `(Tables!$C$810,Tables!$F$810)` → `Tables!$C$810,Tables!$F$810`.

**Rule:** Always handle comma-separated multi-area ranges when reading chart data references. Split, read each part, concatenate.

See: `decx/checker.py` — `_read_chart_range()`

---

## 21. Excel.Quit() Hangs ~60s When COM Chart References Are Alive

**Problem:** `excel_app.Quit()` hangs for ~60 seconds when the Python `with` block still holds live COM references to linked chart shapes (via the `inventory` object). This only manifests with chart-only PPTX files (no OLE objects) because OLE processing masks the timing.

**Root cause:** Python's `with` statement keeps local variables alive until `__exit__` finishes. The `inventory` object holds `(slide, chart_shape)` tuples with live COM pointers. When `Excel.Quit()` runs while these references exist, Excel waits ~60s for COM reference cleanup before actually quitting.

**Debugging evidence:**
- `del inventory` before `with` block exits: 6s total
- Without `del inventory`: 62s total
- `__exit__` called manually (after script-level variables are freed): 1.6s
- Same `__exit__` called via `with` statement: 62s

**Fix (two parts, both required):**
1. `del inventory` inside the `with Session` block before it exits — releases COM chart references
2. Close PowerPoint before Excel in `Session.__exit__` — avoids Excel hanging on chart data source lookups

**Rule:** Always `del inventory` before the `with Session` block ends. And always close PowerPoint before Excel in `__exit__`.

---

## 22. ZIP Pre-Relink Eliminates COM Link Repoint Overhead

**Problem:** `shp.LinkFormat.SourceFullName = new_path` takes ~1s per shape via COM when the path changes. With 186 OLE+chart links, that's ~90-100s just for repointing. For batch runs (26 markets), this adds ~40 minutes of pure relink overhead.

**Why it's slow:** Each `SourceFullName` assignment triggers PowerPoint to validate the new path, connect to Excel internally, and update its link cache — all synchronously via COM.

**Fix:** Rewrite link paths directly in the PPTX zip XML BEFORE COM opens the file. OLE and chart links are stored as external relationships in `.rels` files:
```xml
<Relationship Target="file:///C:\old\data.xlsx!Tables!R1C1" TargetMode="External"/>
```

A simple string replace across all `.rels` files rewrites 186 links in **0.12 seconds**. When COM opens the file, all links already point to the correct Excel — making the COM linker a near-no-op (~0.01s per shape).

**Results:**
- Single file repoint (different Excel): 40s → 21s
- Batch 3 reports: 145s → 64s
- ZIP relink step: 0.12s for 186 links

**Implementation:** `decx/zip_relinker.py` — called in `_run_pairs()` before COM processing. Applied to all update paths (`dx update`, `dx run`).

**Rule:** Always zip-relink before opening with COM. The COM linker still runs as a safety net but becomes a fast no-op.
