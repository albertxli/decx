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

See: `ppt_automation/session.py` — `_auto_dismiss_security_dialog()`

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
