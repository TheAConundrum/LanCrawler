"""Parallel LAN folder crawler (ThreadPool + os.scandir).

Right-click → Run on this file: opens a folder picker, crawls, writes CSV.
"""
from __future__ import annotations

import csv
import datetime as dt
import os
import sys
import time
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

# Exclusive AccDB target beside the workbook: {workbook_dir}\DB\LAN_Search_Index.accdb
ACCDB_DIR_NAME = "DB"
ACCDB_FILE_NAME = "LAN_Search_Index.accdb"
ACCDB_ROW_THRESHOLD = 750_000

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


def accdb_path_for_workbook(workbook: str | Path) -> Path:
    """Relative AccDB beside the workbook: {wb_dir}\\DB\\LAN_Search_Index.accdb."""
    return Path(workbook).resolve().parent / ACCDB_DIR_NAME / ACCDB_FILE_NAME


def _create_empty_accdb(out_path: Path) -> None:
    """Create a blank .accdb via ACE/ADOX (pyodbc cannot create empty AccDB files)."""
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if out_path.exists():
        out_path.unlink()

    providers = (
        "Microsoft.ACE.OLEDB.16.0",
        "Microsoft.ACE.OLEDB.12.0",
    )
    last_err: Exception | None = None
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "Creating AccDB requires pywin32 (win32com). Install: pip install pywin32"
        ) from exc

    for provider in providers:
        try:
            cat = win32com.client.Dispatch("ADOX.Catalog")
            cat.Create(f"Provider={provider};Data Source={out_path};")
            try:
                cat.ActiveConnection.Close()
            except Exception:  # noqa: BLE001
                pass
            return
        except Exception as exc:  # noqa: BLE001
            last_err = exc
            if out_path.exists():
                try:
                    out_path.unlink()
                except OSError:
                    pass

    raise RuntimeError(
        "Could not create AccDB via ADOX. Install Microsoft Access Database Engine (ACE) "
        f"matching Office bitness. Last error: {last_err}"
    ) from last_err


def _accdb_connection_strings(out_path: Path) -> list[str]:
    path = str(Path(out_path).resolve())
    return [
        (
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"DBQ={path};"
        ),
        f"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={path};",
        f"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={path};",
    ]


def write_accdb(rows: list[IndexRow], out_path: Path) -> None:
    """Write IndexRow data to Access .accdb (tblFiles + indexes). Creates file if missing."""
    try:
        import pyodbc  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "Writing AccDB requires pyodbc. Install: pip install pyodbc"
        ) from exc

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Rebuild cleanly: recreate file then bulk insert (avoids stale schema / locks when possible)
    if out_path.exists():
        try:
            out_path.unlink()
        except OSError:
            # File locked — drop table in place instead
            pass

    if not out_path.exists():
        _create_empty_accdb(out_path)

    conn = None
    last_err: Exception | None = None
    for conn_str in _accdb_connection_strings(out_path):
        try:
            conn = pyodbc.connect(conn_str, autocommit=False)
            break
        except Exception as exc:  # noqa: BLE001
            last_err = exc
    if conn is None:
        raise RuntimeError(
            "Could not open AccDB with pyodbc. Install Microsoft Access Database Engine (ACE) "
            f"and the Access ODBC driver. Last error: {last_err}"
        ) from last_err

    try:
        cur = conn.cursor()
        try:
            cur.execute("DROP TABLE tblFiles")
            conn.commit()
        except Exception:  # noqa: BLE001
            conn.rollback()

        cur.execute(
            """
            CREATE TABLE tblFiles (
                FilePath TEXT NOT NULL,
                FileDate TEXT,
                SizeMB DOUBLE NOT NULL,
                EntryType TEXT NOT NULL
            )
            """
        )
        conn.commit()

        insert_sql = (
            "INSERT INTO tblFiles (FilePath, FileDate, SizeMB, EntryType) VALUES (?, ?, ?, ?)"
        )
        batch: list[tuple[str, str | None, float, str]] = []
        batch_size = 500
        for r in rows:
            batch.append((r.path, r.file_date or None, float(r.size_mb), r.entry_type))
            if len(batch) >= batch_size:
                cur.executemany(insert_sql, batch)
                conn.commit()
                batch.clear()
        if batch:
            cur.executemany(insert_sql, batch)
            conn.commit()

        cur.execute("CREATE INDEX ix_tblFiles_path ON tblFiles (FilePath)")
        cur.execute("CREATE INDEX ix_tblFiles_type ON tblFiles (EntryType)")
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


def pick_workbook(
    title: str = "Select workbook to update",
    initial_dir: str | Path | None = None,
) -> str:
    """Open a file dialog for an .xlsm workbook. Returns path or '' if cancelled."""
    import tkinter as tk
    from tkinter import filedialog

    start = Path(initial_dir) if initial_dir else Path(__file__).resolve().parent.parent.parent
    win = tk.Tk()
    win.withdraw()
    try:
        win.attributes("-topmost", True)
        win.lift()
        win.focus_force()
    except tk.TclError:
        pass
    path = filedialog.askopenfilename(
        title=title,
        initialdir=str(start),
        filetypes=[
            ("Excel Macro-Enabled Workbook", "*.xlsm"),
            ("All files", "*.*"),
        ],
    )
    win.destroy()
    return path or ""


def _default_out_dir() -> Path:
    return Path(__file__).resolve().parent.parent.parent / "crawl_output"


def run_interactive(
    *,
    workers: int = 16,
    progress_every: int = 50,
    auto_import_excel: bool = True,
    workbook: str | Path | None = None,
    target_mode: str | None = None,
    use_gui: bool = True,
) -> int:
    """GUI/CLI entry: crawl → Auto/Workbook/AccDB exclusive index write."""
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
    except Exception:
        pass

    print("LAN Search Tool - parallel crawler", flush=True)

    root = ""
    wb_path: Path | None = None
    mode = (target_mode or "Auto").strip() or "Auto"
    if mode not in ("Auto", "Workbook", "AccDB"):
        mode = "Auto"

    if use_gui and workbook is None and target_mode is None:
        try:
            from .crawl_gui import run_crawl_gui
        except ImportError:
            from crawler.crawl_gui import run_crawl_gui  # type: ignore

        gui = run_crawl_gui(initial_workbook=workbook)
        if gui.cancelled:
            print("Cancelled.", flush=True)
            return 0
        root = gui.root
        wb_path = Path(gui.workbook)
        mode = gui.target_mode
    else:
        print("Opening folder picker (check the taskbar if you do not see it)...", flush=True)
        root = pick_folder()
        if not root:
            print("Cancelled - no folder selected.", flush=True)
            return 0

        if workbook is not None and str(workbook).strip():
            wb_path = Path(workbook)
        elif auto_import_excel:
            print("Opening workbook picker (select the .xlsm to update)...", flush=True)
            chosen = pick_workbook(
                initial_dir=Path(__file__).resolve().parent.parent.parent,
                title="Select AccDB / LAN Search workbook",
            )
            if not chosen:
                print("Cancelled - no workbook selected.", flush=True)
                return 0
            wb_path = Path(chosen)
        else:
            print("No workbook selected; AccDB/Excel write skipped.", flush=True)
            wb_path = None

    if wb_path is not None and not wb_path.is_file():
        print(f"Workbook not found: {wb_path}", flush=True)
        return 1

    out_dir = _default_out_dir()
    stamp = time.strftime("%Y%m%d_%H%M%S")
    csv_path = out_dir / f"{stamp}_tblFiles.csv"
    accdb_out = accdb_path_for_workbook(wb_path) if wb_path is not None else None

    drive_map = build_drive_map([root])
    unc = to_unc_path(root, drive_map, "")
    print(f"Root:    {root}", flush=True)
    print(f"UNC:     {unc}", flush=True)
    print(f"Workers: {workers}", flush=True)
    print(f"Mode:    {mode}", flush=True)
    print(f"CSV:     {csv_path}", flush=True)
    if wb_path is not None:
        print(f"Excel:   {wb_path}", flush=True)
    if accdb_out is not None:
        print(f"AccDB:   {accdb_out}", flush=True)
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

    if wb_path is None:
        return 0

    force_accdb = mode == "AccDB" or len(rows) > ACCDB_ROW_THRESHOLD
    if mode == "Workbook" and len(rows) > ACCDB_ROW_THRESHOLD:
        force_accdb = True
        print(
            f"Row count {len(rows):,} exceeds {ACCDB_ROW_THRESHOLD:,}; "
            "forcing AccDB write for performance.",
            flush=True,
        )
    elif mode == "Auto" and len(rows) > ACCDB_ROW_THRESHOLD:
        force_accdb = True
        print(
            f"Auto mode: {len(rows):,} rows > {ACCDB_ROW_THRESHOLD:,} → AccDB.",
            flush=True,
        )

    try:
        if force_accdb:
            assert accdb_out is not None
            print(f"Writing AccDB -> {accdb_out} ...", flush=True)
            write_accdb(rows, accdb_out)
            print(f"Wrote AccDB ({len(rows):,} rows).", flush=True)
            try:
                from .import_to_excel import clear_onboard_tblfiles
            except ImportError:
                from crawler.import_to_excel import clear_onboard_tblfiles  # type: ignore

            print("Clearing onboard Database!tblFiles so AccDB is exclusive...", flush=True)
            clear_onboard_tblfiles(wb_path)
            print(
                f"Index deployed to AccDB:\n  {accdb_out}\n"
                "Onboard sheet database was cleared. Re-open the workbook to search AccDB.",
                flush=True,
            )
        else:
            try:
                from .import_to_excel import import_csv_to_workbook
            except ImportError:
                from crawler.import_to_excel import import_csv_to_workbook  # type: ignore

            print(f"Importing (ReplaceRoot) into workbook {wb_path} ...", flush=True)
            total = import_csv_to_workbook(
                workbook=wb_path,
                csv_path=csv_path,
                mode="ReplaceRoot",
                unc_root=unc,
            )
            print(f"Excel import complete ({total:,} rows in tblFiles).", flush=True)
    except Exception as exc:  # noqa: BLE001
        print(f"Index write failed: {exc}", flush=True)
        print("CSV is still saved.", flush=True)
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
