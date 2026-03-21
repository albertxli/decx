# CLAUDE.md — Project Instructions

## Project

Python rewrite of PowerPoint Excel report automation (from VBA). Uses `win32com` (pywin32) for COM automation.

## Key Commands

```bash
uv run python main.py "report.pptx" --excel "data.xlsx"       # Run automation
uv run pytest tests/ -k "not integration"                       # Unit tests (fast, no COM)
uv run pytest tests/ -k integration                             # Integration tests (needs PPT+Excel)
taskkill /F /IM POWERPNT.EXE & taskkill /F /IM EXCEL.EXE       # Kill zombie COM processes
```

## Critical Rules

1. **READ `GOTCHAS.md` FIRST** before making any changes to COM-related code. It documents every hard-won lesson from debugging pywin32/COM issues. Ignoring it will waste hours.

2. **Track mistakes in `GOTCHAS.md`** — every time you encounter a non-obvious bug, COM quirk, or debugging dead-end during development, document it in `GOTCHAS.md` with the problem, what didn't work, and what does work. This is mandatory, not optional.

3. **VBA is the authoritative reference** — `RUN ALL_Table+Chart_v11.bas` defines all functionality. The old Jupyter notebook (`PPT chart update python.ipynb`) is only useful for COM patterns. Never change logic based on the old Python code.

4. **Use `uv`** for all Python operations (`uv add`, `uv run`, `uv sync`).

5. **Kill zombies before re-running** after a crash — `taskkill /F /IM POWERPNT.EXE`.

## Architecture

```
main.py                    → CLI entry point, batch orchestration
ppt_automation/
  session.py               → COM lifecycle (DispatchEx + dialog auto-dismiss)
  linker.py                → Step 1a: re-point OLE links
  table_updater.py         → Step 1b: populate PPT tables from Excel
  delta_updater.py         → Step 1c: swap delta indicator arrows
  color_coder.py           → Step 1d: _ccst color coding
  chart_updater.py         → Step 2: update chart links
  shape_finder.py          → Token matching, shape discovery
  formatting.py            → Table formatting extract/apply
  utils.py                 → Hex→RGB, R1C1→A1, contrast color
  file_picker.py           → Optional tkinter file dialog (removable)
config.yaml                → All configurable settings
```

## Test Files

`test_files/` contains 1 template PPTX + 3 country Excel files for integration testing.
