"""Small crawl GUI: pick folder (workbook defaults to Lan_Search_Tool.xlsm)."""
from __future__ import annotations

import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

DEFAULT_WORKBOOK_NAME = "Lan_Search_Tool.xlsm"


@dataclass
class CrawlGuiResult:
    root: str
    workbook: str
    target_mode: str  # Auto | Workbook | AccDB
    accdb_fresh: bool = False
    cancelled: bool = False


def _project_root() -> Path:
    return Path(__file__).resolve().parent.parent.parent


def _default_workbook() -> Path:
    return _project_root() / DEFAULT_WORKBOOK_NAME


def ask_accdb_merge_or_fresh(*, taken_on: str) -> str:
    """Ask whether to merge into the existing AccDB or start fresh. Returns add|fresh|cancel."""
    win = tk.Tk()
    win.title("Existing index found")
    win.resizable(False, False)
    try:
        win.attributes("-topmost", True)
        win.lift()
        win.focus_force()
    except tk.TclError:
        pass

    choice = {"value": "cancel"}

    def pick(value: str) -> None:
        choice["value"] = value
        win.destroy()

    frm = ttk.Frame(win, padding=16)
    frm.grid(row=0, column=0, sticky="nsew")
    ttk.Label(
        frm,
        text=(
            f"Do you want to add to the DB taken on {taken_on},\n"
            "or delete it and start fresh?"
        ),
        wraplength=420,
        justify="left",
    ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))

    ttk.Button(frm, text="Add to existing", command=lambda: pick("add")).grid(
        row=1, column=0, padx=4, pady=4, sticky="ew"
    )
    ttk.Button(frm, text="Start fresh", command=lambda: pick("fresh")).grid(
        row=1, column=1, padx=4, pady=4, sticky="ew"
    )
    ttk.Button(frm, text="Cancel", command=lambda: pick("cancel")).grid(
        row=1, column=2, padx=4, pady=4, sticky="ew"
    )
    win.protocol("WM_DELETE_WINDOW", lambda: pick("cancel"))
    win.mainloop()
    return choice["value"]


def run_crawl_gui(
    *,
    initial_workbook: str | Path | None = None,
) -> CrawlGuiResult:
    """Modal dialog. Returns cancelled=True if the user closes without starting."""
    root_win = tk.Tk()
    root_win.title("LAN Search Tool — Crawler")
    root_win.resizable(False, False)
    try:
        root_win.attributes("-topmost", True)
        root_win.lift()
        root_win.focus_force()
    except tk.TclError:
        pass

    folder_var = tk.StringVar(value="")
    wb_default = Path(initial_workbook) if initial_workbook else _default_workbook()
    workbook_var = tk.StringVar(value=str(wb_default) if wb_default.is_file() else "")
    result: dict[str, CrawlGuiResult | None] = {"value": None}
    show_workbook = not wb_default.is_file()

    def browse_folder() -> None:
        path = filedialog.askdirectory(title="Select folder / drive to crawl", mustexist=True)
        if path:
            folder_var.set(path)

    def browse_workbook() -> None:
        start = workbook_var.get().strip()
        initial = Path(start).parent if start else _project_root()
        path = filedialog.askopenfilename(
            title="Select LAN Search workbook (.xlsm)",
            initialdir=str(initial),
            filetypes=[
                ("Excel Macro-Enabled Workbook", "*.xlsm"),
                ("All files", "*.*"),
            ],
        )
        if path:
            workbook_var.set(path)

    def on_start() -> None:
        folder = folder_var.get().strip()
        workbook = workbook_var.get().strip()
        if not folder:
            messagebox.showwarning("Missing folder", "Select a folder or drive to crawl.")
            return
        if not workbook:
            messagebox.showwarning("Missing workbook", "Lan_Search_Tool.xlsm was not found.")
            return
        if not Path(workbook).is_file():
            messagebox.showerror("Workbook not found", f"File not found:\n{workbook}")
            return
        result["value"] = CrawlGuiResult(
            root=folder,
            workbook=workbook,
            target_mode="AccDB",
        )
        root_win.destroy()

    def on_cancel() -> None:
        result["value"] = CrawlGuiResult(
            root="",
            workbook="",
            target_mode="AccDB",
            cancelled=True,
        )
        root_win.destroy()

    padx = 10
    pady = 6
    frm = ttk.Frame(root_win, padding=12)
    frm.grid(row=0, column=0, sticky="nsew")

    ttk.Label(frm, text="Folder / drive to crawl").grid(
        row=0, column=0, sticky="w", padx=padx, pady=pady
    )
    ttk.Entry(frm, textvariable=folder_var, width=64).grid(
        row=1, column=0, sticky="ew", padx=padx, pady=pady
    )
    ttk.Button(frm, text="Browse…", command=browse_folder).grid(
        row=1, column=1, padx=padx, pady=pady
    )

    row = 2
    if show_workbook:
        ttk.Label(frm, text="Target workbook (.xlsm)").grid(
            row=row, column=0, sticky="w", padx=padx, pady=pady
        )
        ttk.Entry(frm, textvariable=workbook_var, width=64).grid(
            row=row + 1, column=0, sticky="ew", padx=padx, pady=pady
        )
        ttk.Button(frm, text="Browse…", command=browse_workbook).grid(
            row=row + 1, column=1, padx=padx, pady=pady
        )
        row = row + 2
    else:
        ttk.Label(frm, text=f"Workbook: {wb_default.name}").grid(
            row=row, column=0, columnspan=2, sticky="w", padx=padx, pady=pady
        )
        row = row + 1

    btn_frm = ttk.Frame(frm)
    btn_frm.grid(row=row, column=0, columnspan=2, sticky="e", padx=padx, pady=pady)
    ttk.Button(btn_frm, text="Cancel", command=on_cancel).pack(side="right", padx=4)
    ttk.Button(btn_frm, text="Start crawl", command=on_start).pack(side="right", padx=4)

    root_win.protocol("WM_DELETE_WINDOW", on_cancel)
    root_win.mainloop()

    value = result["value"]
    if value is None:
        return CrawlGuiResult(root="", workbook="", target_mode="AccDB", cancelled=True)
    return value
