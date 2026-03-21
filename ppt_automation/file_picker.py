"""Optional tkinter file dialog for selecting Excel files.

This module is fully self-contained and can be removed without
affecting any other module. To remove: delete this file and
remove the single `if` block in main.py that calls pick_excel_file().
"""


def pick_excel_file() -> str | None:
    """Show a file dialog to select an Excel file.

    Returns the selected file path, or None if cancelled.
    """
    import ctypes
    import tkinter as tk
    from tkinter import filedialog

    # Make the process DPI-aware so the dialog renders at native resolution
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
    except Exception:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_SYSTEM_DPI_AWARE
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

    root = tk.Tk()
    root.withdraw()  # hide the root window
    root.attributes("-topmost", True)

    file_path = filedialog.askopenfilename(
        title="Select the Excel file for linked data",
        filetypes=[
            ("Excel Files", "*.xlsx *.xlsm *.xls"),
            ("All Files", "*.*"),
        ],
    )

    root.destroy()
    return file_path if file_path else None
