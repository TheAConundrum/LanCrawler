"""Parallel LAN folder crawler (ThreadPool + os.scandir).

Right-click → Run on this file: opens a folder picker, crawls, writes CSV.
"""
from __future__ import annotations

import csv
import datetime as dt
import os
import sqlite3
import sys
import time
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

try:
    from .filters import (
        is_allowed_index_extension,
        is_junk_file_name,
        is_junk_folder_name,
    )
    from .paths import (
        bytes_to_size_mb,
        build_drive_map,
        ensure_crawl_start,
        parent_folder_path,
        path_starts_with_root,
        strip_trailing_slash,
        to_unc_path,
    )
except ImportError:  # python crawl.py (direct / right-click Run)
    _SRC = Path(__file__).resolve().parent.parent
    if str(_SRC) not in sys.path:
        sys.path.insert(0, str(_SRC))
    from crawler.filters import (  # type: ignore
        is_allowed_index_extension,
        is_junk_file_name,
        is_junk_folder_name,
    )
    from crawler.paths import (  # type: ignore
        bytes_to_size_mb,
        build_drive_map,
        ensure_crawl_start,
        parent_folder_path,
        path_starts_with_root,
        strip_trailing_slash,
        to_unc_path,
    )

ENTRY_FILE = "FILE"
ENTRY_FOLDER = "FOLDER"

ProgressCb = Callable[[str], None]


@dataclass
class IndexRow:
    path: str
    file_date: str  # yyyy-mm-dd or ""
    size_mb: float
    entry_type: str


@dataclass
class FolderScanResult:
    folder_path: str
    folder_unc: str
    created_date: str
    local_bytes: int
    file_rows: list[IndexRow] = field(default_factory=list)
    subfolders: list[str] = field(default_factory=list)
    error: str = ""


@dataclass
class CrawlStats:
    folders_scanned: int = 0
    files_indexed: int = 0
    folders_indexed: int = 0
    bytes_all_files: int = 0
    errors: int = 0
    started: float = 0.0
    finished: float = 0.0

    @property
    def elapsed(self) -> float:
        end = self.finished or time.time()
        return max(0.0, end - self.started)


def _date_only_from_timestamp(ts: float | None) -> str:
    if ts is None:
        return ""
    try:
        return dt.datetime.fromtimestamp(ts).date().isoformat()
    except (OverflowError, OSError, ValueError):
        return ""


def _norm_key(path: str) -> str:
    return strip_trailing_slash(path).lower()


def scan_folder(
    folder_path: str,
    drive_map: dict[str, str],
    unc_root_override: str = "",
) -> FolderScanResult:
    """Scan one directory; return indexed files, subfolders, and local byte total."""
    folder_unc = to_unc_path(folder_path, drive_map, unc_root_override)
    result = FolderScanResult(
        folder_path=folder_path,
        folder_unc=folder_unc,
        created_date="",
        local_bytes=0,
    )

    try:
        st_root = os.stat(folder_path, follow_symlinks=False)
        result.created_date = _date_only_from_timestamp(
            getattr(st_root, "st_ctime", None) or st_root.st_mtime
        )
    except OSError:
        pass

    try:
        entries = list(os.scandir(folder_path))
    except OSError as exc:
        result.error = str(exc)
        return result

    for entry in entries:
        name = entry.name
        try:
            is_dir = entry.is_dir(follow_symlinks=False)
        except OSError:
            continue

        if is_dir:
            if is_junk_folder_name(name):
                continue
            result.subfolders.append(entry.path)
            continue

        try:
            st = entry.stat(follow_symlinks=False)
            size = int(st.st_size)
        except OSError:
            size = 0
            st = None

        if size > 0:
            result.local_bytes += size

        if is_junk_file_name(name):
            continue
        if not is_allowed_index_extension(name):
            continue

        unc = to_unc_path(entry.path, drive_map, unc_root_override)
        created = ""
        if st is not None:
            created = _date_only_from_timestamp(getattr(st, "st_ctime", None) or st.st_mtime)

        result.file_rows.append(
            IndexRow(
                path=unc,
                file_date=created,
                size_mb=bytes_to_size_mb(size),
                entry_type=ENTRY_FILE,
            )
        )

    return result


def _bubble_folder_bytes(local_bytes: dict[str, int], crawl_root_unc: str) -> dict[str, int]:
    totals: dict[str, int] = {}
    root = strip_trailing_slash(crawl_root_unc)

    for folder_unc, nbytes in local_bytes.items():
        if nbytes <= 0:
            continue
        p = strip_trailing_slash(folder_unc)
        while p:
            totals[p] = totals.get(p, 0) + nbytes
            if p.lower() == root.lower():
                break
            parent = parent_folder_path(p)
            if not parent or parent.lower() == p.lower():
                break
            if parent.lower() != root.lower() and not path_starts_with_root(parent, root):
                break
            p = parent
    return totals


def crawl_parallel(
    root_path: str,
    *,
    workers: int = 16,
    drive_map: dict[str, str] | None = None,
    unc_root_override: str = "",
    progress_every: int = 50,
    on_progress: ProgressCb | None = None,
) -> tuple[list[IndexRow], CrawlStats]:
    """Crawl root_path with a thread pool. Returns (tblFiles rows, stats)."""
    start = ensure_crawl_start(root_path)
    start_s = str(start)
    drive_map = drive_map if drive_map is not None else {}
    crawl_root_unc = to_unc_path(start_s, drive_map, unc_root_override)

    stats = CrawlStats(started=time.time())
    file_rows: list[IndexRow] = []
    folder_meta: dict[str, str] = {}
    local_bytes: dict[str, int] = {}

    def emit(msg: str) -> None:
        if on_progress:
            on_progress(msg)

    seen: set[str] = set()
    n_workers = max(1, int(workers))

    with ThreadPoolExecutor(max_workers=n_workers, thread_name_prefix="lan-crawl") as pool:
        pending = set()
        seen.add(_norm_key(start_s))
        pending.add(pool.submit(scan_folder, start_s, drive_map, unc_root_override))

        while pending:
            done, pending = wait(pending, return_when=FIRST_COMPLETED)
            for fut in done:
                try:
                    res: FolderScanResult = fut.result()
                except Exception as exc:  # noqa: BLE001 — isolate worker failures
                    stats.errors += 1
                    emit(f"worker error: {exc}")
                    continue

                if res.error:
                    stats.errors += 1
                    emit(f"skip {res.folder_path}: {res.error}")
                else:
                    stats.folders_scanned += 1
                    if res.folder_unc:
                        folder_meta[res.folder_unc] = res.created_date
                        local_bytes[res.folder_unc] = local_bytes.get(res.folder_unc, 0) + res.local_bytes
                    file_rows.extend(res.file_rows)

                    if stats.folders_scanned % progress_every == 0:
                        emit(
                            f"folders={stats.folders_scanned:,} "
                            f"files={len(file_rows):,} "
                            f"pending={len(pending):,} "
                            f"errors={stats.errors:,}"
                        )

                for sub in res.subfolders:
                    key = _norm_key(sub)
                    if key in seen:
                        continue
                    # Skip junk by name even if parent listing included it
                    if is_junk_folder_name(Path(sub).name):
                        continue
                    seen.add(key)
                    pending.add(pool.submit(scan_folder, sub, drive_map, unc_root_override))

    totals = _bubble_folder_bytes(local_bytes, crawl_root_unc)

    rows: list[IndexRow] = list(file_rows)
    for unc, created in folder_meta.items():
        rows.append(
            IndexRow(
                path=unc,
                file_date=created,
                size_mb=bytes_to_size_mb(totals.get(unc, 0)),
                entry_type=ENTRY_FOLDER,
            )
        )

    stats.files_indexed = sum(1 for r in rows if r.entry_type == ENTRY_FILE)
    stats.folders_indexed = sum(1 for r in rows if r.entry_type == ENTRY_FOLDER)
    stats.bytes_all_files = sum(local_bytes.values())
    stats.finished = time.time()
    emit(
        f"done files={stats.files_indexed:,} folders={stats.folders_indexed:,} "
        f"errors={stats.errors:,} elapsed={stats.elapsed:.1f}s "
        f"disk_bytes={stats.bytes_all_files:,}"
    )
    return rows, stats


def write_csv(rows: list[IndexRow], out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["FilePath", "FileDate", "SizeMB", "EntryType"])
        for r in rows:
            w.writerow([r.path, r.file_date, f"{r.size_mb:.2f}", r.entry_type])


def write_sqlite(rows: list[IndexRow], out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if out_path.exists():
        out_path.unlink()
    conn = sqlite3.connect(str(out_path))
    try:
        conn.execute(
            """
            CREATE TABLE tblFiles (
                FilePath TEXT NOT NULL,
                FileDate TEXT,
                SizeMB REAL NOT NULL,
                EntryType TEXT NOT NULL
            )
            """
        )
        conn.executemany(
            "INSERT INTO tblFiles (FilePath, FileDate, SizeMB, EntryType) VALUES (?, ?, ?, ?)",
            [(r.path, r.file_date or None, r.size_mb, r.entry_type) for r in rows],
        )
        conn.execute("CREATE INDEX ix_tblFiles_path ON tblFiles(FilePath)")
        conn.execute("CREATE INDEX ix_tblFiles_type ON tblFiles(EntryType)")
        conn.commit()
    finally:
        conn.close()


def pick_folder(title: str = "Select a folder to crawl (LAN Search Tool)") -> str:
    """
    Windows folder browser.
    Prefers WinForms (shows in front of Cursor); falls back to tkinter.
    """
    # WinForms via PowerShell - more reliable than tkinter when launched from an IDE
    if os.name == "nt":
        safe_title = title.replace("'", "''")
        ps = (
            "Add-Type -AssemblyName System.Windows.Forms; "
            "$d = New-Object System.Windows.Forms.FolderBrowserDialog; "
            f"$d.Description = '{safe_title}'; "
            "$d.ShowNewFolderButton = $true; "
            "$r = $d.ShowDialog(); "
            "if ($r -eq [System.Windows.Forms.DialogResult]::OK) { "
            "  [Console]::Out.Write($d.SelectedPath) "
            "}"
        )
        try:
            import subprocess

            completed = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-STA",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    ps,
                ],
                capture_output=True,
                text=True,
                timeout=600,
                check=False,
            )
            path = (completed.stdout or "").strip()
            if path:
                return path
            # Empty stdout with return code 0 usually means Cancel
            if completed.returncode == 0 and not (completed.stderr or "").strip():
                return ""
            # Fall through to tkinter if PowerShell failed oddly
            if completed.stderr:
                print(f"(folder dialog fallback: {completed.stderr.strip()[:200]})", flush=True)
        except Exception as exc:  # noqa: BLE001
            print(f"(folder dialog fallback: {exc})", flush=True)

    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
        root.lift()
        root.focus_force()
    except tk.TclError:
        pass
    path = filedialog.askdirectory(title=title, mustexist=True)
    root.destroy()
    return path or ""


def _default_out_dir() -> Path:
    return Path(__file__).resolve().parent.parent.parent / "crawl_output"


def _default_workbook() -> Path | None:
    wb = Path(__file__).resolve().parent.parent.parent / "Blank_LAN_Crawler Tool.xlsm"
    return wb if wb.exists() else None


def run_interactive(
    *,
    workers: int = 16,
    progress_every: int = 50,
    auto_import_excel: bool = True,
) -> int:
    """Folder-picker entry: crawl → CSV → auto-import Excel (no prompts after pick)."""
    # Force UTF-8-friendly console where possible; avoid fancy punctuation either way
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
    except Exception:
        pass

    print("LAN Search Tool - parallel crawler", flush=True)
    print("Opening folder picker (check the taskbar if you do not see it)...", flush=True)
    root = pick_folder()
    if not root:
        print("Cancelled - no folder selected.", flush=True)
        return 0

    out_dir = _default_out_dir()
    stamp = time.strftime("%Y%m%d_%H%M%S")
    csv_path = out_dir / f"{stamp}_tblFiles.csv"

    drive_map = build_drive_map([root])
    unc = to_unc_path(root, drive_map, "")
    print(f"Root:    {root}", flush=True)
    print(f"UNC:     {unc}", flush=True)
    print(f"Workers: {workers}", flush=True)
    print(f"CSV:     {csv_path}", flush=True)
    print("---", flush=True)

    def on_progress(msg: str) -> None:
        print(msg, flush=True)

    t0 = time.time()
    rows, stats = crawl_parallel(
        root,
        workers=workers,
        drive_map=drive_map,
        progress_every=progress_every,
        on_progress=on_progress,
    )
    write_csv(rows, csv_path)
    print("---", flush=True)
    print(f"Wrote {len(rows):,} rows -> {csv_path}", flush=True)
    print(
        f"files={stats.files_indexed:,} folders={stats.folders_indexed:,} "
        f"errors={stats.errors:,} disk~{bytes_to_size_mb(stats.bytes_all_files):,.2f} MB "
        f"in {time.time() - t0:.1f}s",
        flush=True,
    )

    if not auto_import_excel:
        return 0

    wb = _default_workbook()
    if wb is None:
        print("No Blank_LAN_Crawler Tool.xlsm found next to crawl_output — CSV only.", flush=True)
        return 0

    try:
        try:
            from .import_to_excel import import_csv_to_workbook
        except ImportError:
            from crawler.import_to_excel import import_csv_to_workbook  # type: ignore

        print(f"Importing (ReplaceRoot) into {wb} ...", flush=True)
        total = import_csv_to_workbook(
            workbook=wb,
            csv_path=csv_path,
            mode="ReplaceRoot",
            unc_root=unc,
        )
        print(f"Excel import complete ({total:,} rows in tblFiles).", flush=True)
    except Exception as exc:  # noqa: BLE001
        print(f"Excel import failed: {exc}", flush=True)
        print("CSV is still saved - you can import later with:", flush=True)
        print(f'  python -m crawler "{root}" --import-excel "{wb}"', flush=True)
        return 1

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(run_interactive())
    except SystemExit:
        raise
    except Exception as exc:  # noqa: BLE001
        import traceback

        print("CRASH:", exc, flush=True)
        traceback.print_exc()
        raise SystemExit(1) from exc
