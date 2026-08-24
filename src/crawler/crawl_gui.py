"""Crawl setup + live progress window."""
from __future__ import annotations

import queue
import threading
import time
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Callable

from crawler.crawl import CrawlStats

DEFAULT_WORKBOOK_NAME = "Lan_Search_Tool.xlsm"

# quick = incremental update; full = recrawl this folder; wipe = delete AccDB first
MODE_QUICK = "quick"
MODE_FULL = "full"
MODE_WIPE = "wipe"


@dataclass
class CrawlGuiResult:
    root: str
    workbook: str
    target_mode: str  # Auto | Workbook | AccDB
    workers: int = 16
    crawl_mode: str = MODE_QUICK
    accdb_fresh: bool = False
    incremental: bool = True
    cancelled: bool = False


def _project_root() -> Path:
    return Path(__file__).resolve().parent.parent.parent


def _default_workbook() -> Path:
    return _project_root() / DEFAULT_WORKBOOK_NAME


def _fmt_hms(seconds: float) -> str:
    sec = max(0, int(seconds))
    h, rem = divmod(sec, 3600)
    m, s = divmod(rem, 60)
    if h:
        return f"{h}:{m:02d}:{s:02d}"
    return f"{m}:{s:02d}"


def run_crawl_gui(
    *,
    initial_workbook: str | Path | None = None,
    has_existing_index: bool | None = None,
) -> CrawlGuiResult:
    """Modal setup dialog. Returns cancelled=True if the user closes without starting."""
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
    workers_var = tk.IntVar(value=16)
    mode_var = tk.StringVar(value=MODE_QUICK if has_existing_index else MODE_FULL)
    result: dict[str, CrawlGuiResult | None] = {"value": None}
    show_workbook = not wb_default.is_file()
    existing = bool(has_existing_index) if has_existing_index is not None else False
    if not existing and wb_default.is_file():
        try:
            from crawler.crawl import existing_accdb_for_workbook

            existing = existing_accdb_for_workbook(wb_default, warn=False) is not None
            if existing:
                mode_var.set(MODE_QUICK)
        except Exception:
            existing = False

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
        try:
            workers = int(workers_var.get())
        except (TypeError, ValueError, tk.TclError):
            workers = 16
        workers = max(1, min(48, workers))
        mode = mode_var.get().strip() or MODE_FULL
        if not existing:
            mode = MODE_FULL
        result["value"] = CrawlGuiResult(
            root=folder,
            workbook=workbook,
            target_mode="AccDB",
            workers=workers,
            crawl_mode=mode,
            accdb_fresh=mode == MODE_WIPE,
            incremental=mode == MODE_QUICK,
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
    frm = ttk.Frame(root_win, padding=14)
    frm.grid(row=0, column=0, sticky="nsew")
    frm.columnconfigure(0, weight=1)

    ttk.Label(frm, text="Folder / drive to crawl").grid(
        row=0, column=0, sticky="w", padx=padx, pady=(0, 2)
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
            row=row, column=0, sticky="w", padx=padx, pady=(8, 2)
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

    wrk = ttk.Frame(frm)
    wrk.grid(row=row, column=0, columnspan=2, sticky="w", padx=padx, pady=(8, 4))
    ttk.Label(wrk, text="Parallel workers").pack(side="left")
    ttk.Spinbox(wrk, from_=4, to=48, textvariable=workers_var, width=6).pack(
        side="left", padx=(8, 0)
    )
    ttk.Label(wrk, text="(16 is safe; try 24–32 if the file server is idle)").pack(
        side="left", padx=(10, 0)
    )
    row += 1

    ttk.Label(frm, text="Update mode").grid(
        row=row, column=0, sticky="w", padx=padx, pady=(10, 2)
    )
    row += 1
    modes = ttk.Frame(frm)
    modes.grid(row=row, column=0, columnspan=2, sticky="w", padx=padx, pady=pady)
    ttk.Radiobutton(
        modes,
        text="Quick update — only folders with new or changed files (recommended)",
        variable=mode_var,
        value=MODE_QUICK,
    ).grid(row=0, column=0, sticky="w")
    ttk.Radiobutton(
        modes,
        text="Full recrawl — re-list every folder under this path (other drives kept)",
        variable=mode_var,
        value=MODE_FULL,
    ).grid(row=1, column=0, sticky="w")
    ttk.Radiobutton(
        modes,
        text="Start fresh — erase the entire index, then crawl this folder only",
        variable=mode_var,
        value=MODE_WIPE,
    ).grid(row=2, column=0, sticky="w")
    if not existing:
        mode_var.set(MODE_FULL)
        for child in modes.winfo_children()[:1]:
            child.configure(state="disabled")
        ttk.Label(
            frm,
            text="No existing AccDB yet — first crawl is always a full listing.",
            foreground="#555",
        ).grid(row=row + 1, column=0, columnspan=2, sticky="w", padx=padx, pady=(0, 6))
        row += 1
    row += 1

    hint = ttk.Label(
        frm,
        text=(
            "Quick update: a resaved file is picked up even if the folder date did not change.\n"
            "A file added in a subfolder is picked up by walking that subfolder."
        ),
        foreground="#333",
        justify="left",
    )
    hint.grid(row=row, column=0, columnspan=2, sticky="w", padx=padx, pady=(4, 8))
    row += 1

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


def run_progress_window(
    *,
    root_label: str,
    mode_label: str,
    work: Callable[[threading.Event, Callable[[CrawlStats], None], Callable[[str], None]], None],
) -> bool:
    """
    Show live counters while `work` runs on a background thread.

    work(cancel_event, on_stats, on_log)
    Returns False if the user cancelled before completion.
    """
    win = tk.Tk()
    win.title("LAN Search Tool — Crawling")
    win.minsize(560, 420)
    try:
        win.attributes("-topmost", True)
        win.lift()
    except tk.TclError:
        pass

    cancel_event = threading.Event()
    msg_q: queue.Queue = queue.Queue()
    started = time.time()
    finished = {"ok": False, "cancelled": False, "message": ""}

    frm = ttk.Frame(win, padding=14)
    frm.pack(fill="both", expand=True)
    frm.columnconfigure(1, weight=1)

    ttk.Label(frm, text="Root").grid(row=0, column=0, sticky="w")
    ttk.Label(frm, text=root_label, wraplength=480).grid(row=0, column=1, sticky="w", padx=(8, 0))
    ttk.Label(frm, text="Mode").grid(row=1, column=0, sticky="w", pady=(4, 0))
    ttk.Label(frm, text=mode_label).grid(row=1, column=1, sticky="w", padx=(8, 0), pady=(4, 0))

    pbar = ttk.Progressbar(frm, mode="determinate", maximum=100)
    pbar.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(14, 8))

    stats_frm = ttk.Frame(frm)
    stats_frm.grid(row=3, column=0, columnspan=2, sticky="ew")
    labels = [
        ("folders", "Folders done"),
        ("scan", "Full listings"),
        ("quick", "Quick (unchanged)"),
        ("files", "Files"),
        ("pending", "Pending"),
        ("errors", "Errors"),
        ("elapsed", "Elapsed"),
    ]
    vars_map: dict[str, tk.StringVar] = {}
    for i, (key, title) in enumerate(labels):
        r, c = divmod(i, 4)
        cell = ttk.Frame(stats_frm, padding=6)
        cell.grid(row=r, column=c, sticky="w", padx=4, pady=2)
        ttk.Label(cell, text=title, foreground="#555").pack(anchor="w")
        var = tk.StringVar(value="0")
        vars_map[key] = var
        ttk.Label(cell, textvariable=var, font=("Segoe UI", 14, "bold")).pack(anchor="w")

    status_var = tk.StringVar(value="Starting…")
    ttk.Label(frm, textvariable=status_var).grid(
        row=4, column=0, columnspan=2, sticky="w", pady=(10, 4)
    )

    ttk.Label(frm, text="Log").grid(row=5, column=0, columnspan=2, sticky="w")
    log = tk.Text(frm, height=10, wrap="word", state="disabled")
    log.grid(row=6, column=0, columnspan=2, sticky="nsew", pady=(2, 8))
    frm.rowconfigure(6, weight=1)
    scroll = ttk.Scrollbar(frm, command=log.yview)
    scroll.grid(row=6, column=2, sticky="ns")
    log.configure(yscrollcommand=scroll.set)

    btn_frm = ttk.Frame(frm)
    btn_frm.grid(row=7, column=0, columnspan=2, sticky="e")
    close_btn = ttk.Button(btn_frm, text="Cancel crawl")

    def append_log(text: str) -> None:
        log.configure(state="normal")
        log.insert("end", text + "\n")
        log.see("end")
        log.configure(state="disabled")

    def apply_stats(st: CrawlStats, status: str | None = None) -> None:
        vars_map["folders"].set(f"{st.folders_done:,}")
        vars_map["scan"].set(f"{st.folders_scanned:,}")
        vars_map["quick"].set(f"{st.folders_quick:,}")
        vars_map["files"].set(f"{st.files_indexed:,}")
        vars_map["pending"].set(f"{st.pending:,}")
        vars_map["errors"].set(f"{st.errors:,}")
        vars_map["elapsed"].set(_fmt_hms(st.elapsed if st.started else time.time() - started))
        done = st.folders_done
        pending = st.pending
        total = done + pending
        if total > 0:
            pbar.configure(mode="determinate", maximum=max(total, 1))
            pbar["value"] = done
        else:
            pbar.configure(mode="indeterminate")
        if status:
            status_var.set(status)

    def on_close_request() -> None:
        if not finished["ok"] and not finished["cancelled"] and not cancel_event.is_set():
            if not messagebox.askyesno("Cancel crawl", "Stop the crawl? Nothing will be written."):
                return
            cancel_event.set()
            status_var.set("Cancelling…")
            return
        win.destroy()

    close_btn.configure(command=on_close_request)
    close_btn.pack(side="right")

    def on_stats(st: CrawlStats) -> None:
        msg_q.put(("stats", st, None))

    def on_log(msg: str) -> None:
        msg_q.put(("log", msg))

    def worker() -> None:
        try:
            work(cancel_event, on_stats, on_log)
            if cancel_event.is_set():
                msg_q.put(("done", False, "Cancelled. Index was not updated."))
            else:
                msg_q.put(("done", True, "Finished."))
        except Exception as exc:  # noqa: BLE001
            msg_q.put(("done", False, f"Failed: {exc}"))

    def poll() -> None:
        try:
            while True:
                kind, *rest = msg_q.get_nowait()
                if kind == "stats":
                    apply_stats(rest[0], rest[1])
                elif kind == "log":
                    text = str(rest[0])
                    if not text.startswith("folders="):
                        append_log(text)
                    status_var.set(text[:160])
                elif kind == "status":
                    status_var.set(str(rest[0]))
                elif kind == "done":
                    ok, message = rest[0], rest[1]
                    finished["ok"] = bool(ok)
                    finished["cancelled"] = cancel_event.is_set() or not ok
                    finished["message"] = message
                    status_var.set(message)
                    append_log(message)
                    pbar.stop()
                    if ok:
                        pbar.configure(mode="determinate", maximum=100)
                        pbar["value"] = 100
                    close_btn.configure(text="Close")
                    return
        except queue.Empty:
            pass
        if not finished["ok"] and not finished["message"]:
            vars_map["elapsed"].set(_fmt_hms(time.time() - started))
        win.after(100, poll)

    threading.Thread(target=worker, daemon=True, name="lan-crawl-ui").start()
    win.protocol("WM_DELETE_WINDOW", on_close_request)
    win.after(100, poll)
    win.mainloop()
    return bool(finished["ok"])


# Kept for older call sites
def ask_accdb_merge_or_fresh(*, taken_on: str) -> str:
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
