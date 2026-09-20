"""
Shift report filler: pops up a small window asking for the check date and
examiner name, then updates those two fields in the Word document and saves
it with a filename matching the date.

Requires: pip install python-docx

Usage: python fill_report.py
(Optionally edit TEMPLATE_PATH below if your template file has a different name.)
"""

import tkinter as tk
from tkinter import messagebox
from datetime import datetime
from pathlib import Path

import docx

OUTPUT_DIR = Path("phonechecks")  # where to save the filled-in copy
OUTPUT_DIR.mkdir(exist_ok=True)
TEMPLATE_PATH = OUTPUT_DIR / "19_09_26.docx"  # the file to use as the starting template


def update_report(date_str: str, name_str: str) -> Path:
    doc = docx.Document(TEMPLATE_PATH)
    table = doc.tables[0]
    table.rows[0].cells[1].text = date_str  # תאריך בדיקה
    table.rows[0].cells[3].text = name_str  # שם הבודק

    safe_date = date_str.replace(".", "_").replace("/", "_")
    out_path = OUTPUT_DIR / f"{safe_date}.docx"
    doc.save(out_path)
    return out_path


def on_submit():
    date_str = date_entry.get().strip()
    name_str = name_entry.get().strip()

    if not date_str or not name_str:
        messagebox.showwarning("Missing info", "Please fill in both the date and the name.")
        return

    if not TEMPLATE_PATH.exists():
        messagebox.showerror("Template not found", f"Could not find {TEMPLATE_PATH.resolve()}")
        return

    try:
        out_path = update_report(date_str, name_str)
        messagebox.showinfo("Saved", f"Saved: {out_path.resolve()}")
        root.destroy()
    except Exception as e:
        messagebox.showerror("Error", str(e))


root = tk.Tk()
root.title("Shift Report")
root.geometry("320x150")

tk.Label(root, text="Check date (DD.MM.YY):").pack(pady=(15, 0))
date_entry = tk.Entry(root, justify="center")
date_entry.insert(0, datetime.now().strftime("%d.%m.%y"))
date_entry.pack(pady=5)

tk.Label(root, text="Examiner name:").pack()
name_entry = tk.Entry(root, justify="center")
name_entry.pack(pady=5)

tk.Button(root, text="Save", command=on_submit, width=12).pack(pady=15)

root.mainloop()