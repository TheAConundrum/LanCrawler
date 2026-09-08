"""Parallel LAN folder crawler (ThreadPool + os.scandir).

Right-click → Run on this file: opens a folder picker, crawls, writes AccDB.
"""
from __future__ import annotations

import csv
import datetime as dt
import os
import shutil
import struct
import sys
import tempfile
import threading
import time
from collections import deque
from concurrent.futures import FIRST_COMPLETED, Future, ThreadPoolExecutor, wait
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

# Exclusive AccDB beside the workbook:
#   newest DB\SearchIndex-{M}-{D}-{YYYY}.accdb (filename = snapshot birth date; crawls merge in)
#   else LAN_Search_Index.accdb
#   else create SearchIndex-{today}.accdb
ACCDB_DIR_NAME = "DB"
ACCDB_FILE_PREFIX = "SearchIndex"
ACCDB_FILE_NAME_LEGACY = "LAN_Search_Index.accdb"
ACCDB_FILE_NAME = ACCDB_FILE_NAME_LEGACY  # kept for import compatibility; dated names are preferred
DEFAULT_WORKBOOK_NAME = "Lan_Search_Tool.xlsm"
ACCDB_ROW_THRESHOLD = 750_000

try:
    from .filters import (
        is_allowed_index_extension,
        is_junk_file_name,
        is_junk_folder_name,
    )
    from .index_state import (
        FolderMetaRow,
        FolderSnapshot,
        PreviousIndex,
        build_previous_index,
        folder_mtime_unchanged,
        row_size_bytes,
    )
    from .paths import (
        bytes_to_size_mb,
        build_drive_map,
        ensure_crawl_start,
        parent_folder_path,
        path_starts_with_root,
        strip_trailing_slash,
        to_unc_path,
        MIN_SIZE_MB,
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
    from crawler.index_state import (  # type: ignore
        FolderMetaRow,
        FolderSnapshot,
        PreviousIndex,
        build_previous_index,
        folder_mtime_unchanged,
        row_size_bytes,
    )
    from crawler.paths import (  # type: ignore
        bytes_to_size_mb,
        build_drive_map,
        ensure_crawl_start,
        parent_folder_path,
        path_starts_with_root,
        strip_trailing_slash,
        to_unc_path,
        MIN_SIZE_MB,
    )

ENTRY_FILE = "FILE"
ENTRY_FOLDER = "FOLDER"

ProgressCb = Callable[[str], None]
StatsCb = Callable[["CrawlStats"], None]


@dataclass
class IndexRow:
    path: str
    file_date: str  # yyyy-mm-dd (files: last modified; folders: created)
    size_mb: float
    entry_type: str
    size_unknown: bool = False  # True only when the crawl stat failed
    mtime_ts: float = 0.0  # file mtime or folder dir mtime
    size_bytes: int = 0  # file bytes or folder local (non-recursive) bytes


@dataclass
class FolderScanResult:
    folder_path: str
    folder_unc: str
    created_date: str
    local_bytes: int
    file_rows: list[IndexRow] = field(default_factory=list)
    subfolders: list[str] = field(default_factory=list)
    error: str = ""
    dir_mtime: float = 0.0
    skipped: bool = False  # True = dir mtime unchanged; restated known files only


@dataclass
class CrawlStats:
    folders_scanned: int = 0
    folders_quick: int = 0
    files_indexed: int = 0
    folders_indexed: int = 0
    bytes_all_files: int = 0
    errors: int = 0
    pending: int = 0
    cancelled: bool = False
    incremental: bool = False
    started: float = 0.0
    finished: float = 0.0

    @property
    def folders_done(self) -> int:
        return self.folders_scanned + self.folders_quick

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
        result.dir_mtime = float(st_root.st_mtime)
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
            size_unknown = False
        except OSError:
            size = 0
            st = None
            size_unknown = True

        if size > 0:
            result.local_bytes += size

        if is_junk_file_name(name):
            continue
        if not is_allowed_index_extension(name):
            continue

        unc = to_unc_path(entry.path, drive_map, unc_root_override)
        mtime_ts = float(st.st_mtime) if st is not None else 0.0
        modified = _date_only_from_timestamp(mtime_ts if mtime_ts else None)

        result.file_rows.append(
            IndexRow(
                path=unc,
                file_date=modified,
                size_mb=bytes_to_size_mb(size),
                entry_type=ENTRY_FILE,
                size_unknown=size_unknown,
                mtime_ts=mtime_ts,
                size_bytes=int(size),
            )
        )

    return result


def restat_known_files(
    prev: FolderSnapshot,
) -> tuple[list[IndexRow], int]:
    """Re-stat indexed files in an unchanged folder (catches in-place resaves)."""
    new_rows: list[IndexRow] = []
    old_indexed = 0
    new_indexed = 0
    for snap in prev.files:
        old_b = row_size_bytes(snap.size_bytes, snap.size_mb)
        old_indexed += old_b
        try:
            st = os.stat(snap.path, follow_symlinks=False)
            size = int(st.st_size)
            mtime_ts = float(st.st_mtime)
            new_rows.append(
                IndexRow(
                    path=snap.path,
                    file_date=_date_only_from_timestamp(mtime_ts),
                    size_mb=bytes_to_size_mb(size),
                    entry_type=ENTRY_FILE,
                    mtime_ts=mtime_ts,
                    size_bytes=size,
                )
            )
            new_indexed += size
        except OSError:
            continue
    local_bytes = int(prev.local_bytes) - old_indexed + new_indexed
    if local_bytes < 0:
        local_bytes = new_indexed
    return new_rows, local_bytes


def visit_folder(
    folder_path: str,
    drive_map: dict[str, str],
    unc_root_override: str,
    prev: FolderSnapshot | None,
    incremental: bool,
) -> FolderScanResult:
    """Full scandir, or skip listing when dir mtime is unchanged."""
    folder_unc = to_unc_path(folder_path, drive_map, unc_root_override)
    dir_mtime = 0.0
    created = ""
    try:
        st_root = os.stat(folder_path, follow_symlinks=False)
        dir_mtime = float(st_root.st_mtime)
        created = _date_only_from_timestamp(
            getattr(st_root, "st_ctime", None) or st_root.st_mtime
        )
    except OSError as exc:
        return FolderScanResult(
            folder_path=folder_path,
            folder_unc=folder_unc,
            created_date="",
            local_bytes=0,
            error=str(exc),
        )

    can_skip = (
        incremental
        and prev is not None
        and folder_mtime_unchanged(dir_mtime, prev.dir_mtime)
    )
    if can_skip:
        assert prev is not None
        file_rows, local_bytes = restat_known_files(prev)
        return FolderScanResult(
            folder_path=folder_path,
            folder_unc=folder_unc,
            created_date=prev.created_date or created,
            local_bytes=local_bytes,
            file_rows=file_rows,
            subfolders=list(prev.child_uncs),
            dir_mtime=dir_mtime,
            skipped=True,
        )

    result = scan_folder(folder_path, drive_map, unc_root_override)
    if not result.dir_mtime:
        result.dir_mtime = dir_mtime
    if not result.created_date:
        result.created_date = created
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
    on_stats: StatsCb | None = None,
    incremental: bool = False,
    previous: PreviousIndex | None = None,
    cancel_event: threading.Event | None = None,
) -> tuple[list[IndexRow], CrawlStats]:
    """Crawl root_path with a thread pool. Returns (tblFiles rows, stats).

    incremental=True plus a PreviousIndex with folder timestamps: skip scandir on
    unchanged folders, re-stat known files, and still walk known child folders.
    """
    start = ensure_crawl_start(root_path)
    start_s = str(start)
    drive_map = drive_map if drive_map is not None else {}
    crawl_root_unc = to_unc_path(start_s, drive_map, unc_root_override)
    use_incremental = bool(incremental and previous is not None and previous.has_meta)

    stats = CrawlStats(started=time.time(), incremental=use_incremental)
    file_rows: list[IndexRow] = []
    folder_meta: dict[str, str] = {}
    folder_mtime: dict[str, float] = {}
    local_bytes: dict[str, int] = {}

    def emit(msg: str) -> None:
        if on_progress:
            on_progress(msg)

    def push_stats() -> None:
        stats.files_indexed = len(file_rows)
        stats.folders_indexed = len(folder_meta)
        if on_stats:
            on_stats(stats)

    def mark_seen(path: str, seen: set[str]) -> bool:
        key = _norm_key(path)
        unc_key = _norm_key(to_unc_path(path, drive_map, unc_root_override))
        if key in seen or unc_key in seen:
            return False
        seen.add(key)
        if unc_key:
            seen.add(unc_key)
        return True

    seen: set[str] = set()
    n_workers = max(1, int(workers))
    max_outstanding = max(n_workers * 8, 32)
    queued: deque[str] = deque()
    in_flight: dict[Future, str] = {}

    mark_seen(start_s, seen)
    queued.append(start_s)

    def submit_more(pool: ThreadPoolExecutor) -> None:
        while queued and len(in_flight) < max_outstanding:
            if cancel_event is not None and cancel_event.is_set():
                return
            path = queued.popleft()
            prev_snap = previous.get(to_unc_path(path, drive_map, unc_root_override)) if previous else None
            fut = pool.submit(
                visit_folder,
                path,
                drive_map,
                unc_root_override,
                prev_snap,
                use_incremental,
            )
            in_flight[fut] = path

    with ThreadPoolExecutor(max_workers=n_workers, thread_name_prefix="lan-crawl") as pool:
        submit_more(pool)
        while in_flight or queued:
            if cancel_event is not None and cancel_event.is_set() and not in_flight:
                break
            if not in_flight:
                submit_more(pool)
                if not in_flight:
                    break
            done, _still = wait(list(in_flight.keys()), return_when=FIRST_COMPLETED)
            progressed = False
            for fut in done:
                in_flight.pop(fut, None)
                try:
                    res: FolderScanResult = fut.result()
                except Exception as exc:  # noqa: BLE001 — isolate worker failures
                    stats.errors += 1
                    emit(f"worker error: {exc}")
                    continue

                if res.error:
                    stats.errors += 1
                    emit(f"skip {res.folder_path}: {res.error}")
                elif res.skipped:
                    stats.folders_quick += 1
                    progressed = True
                    if res.folder_unc:
                        folder_meta[res.folder_unc] = res.created_date
                        folder_mtime[res.folder_unc] = res.dir_mtime
                        local_bytes[res.folder_unc] = res.local_bytes
                    file_rows.extend(res.file_rows)
                else:
                    stats.folders_scanned += 1
                    progressed = True
                    if res.folder_unc:
                        folder_meta[res.folder_unc] = res.created_date
                        folder_mtime[res.folder_unc] = res.dir_mtime
                        local_bytes[res.folder_unc] = local_bytes.get(res.folder_unc, 0) + res.local_bytes
                    file_rows.extend(res.file_rows)

                for sub in res.subfolders:
                    if is_junk_folder_name(Path(sub).name):
                        continue
                    if mark_seen(sub, seen):
                        queued.append(sub)

            stats.pending = len(in_flight) + len(queued)
            if progressed and stats.folders_done > 0 and stats.folders_done % progress_every == 0:
                emit(
                    f"folders={stats.folders_done:,} "
                    f"scan={stats.folders_scanned:,} "
                    f"quick={stats.folders_quick:,} "
                    f"files={len(file_rows):,} "
                    f"pending={stats.pending:,} "
                    f"errors={stats.errors:,}"
                )
            push_stats()
            if cancel_event is not None and cancel_event.is_set():
                queued.clear()
                stats.cancelled = True
                emit("cancelled — stopping after in-flight folders")
                # Drain remaining futures so the pool can exit cleanly
                continue
            submit_more(pool)

    if stats.cancelled:
        stats.files_indexed = len(file_rows)
        stats.folders_indexed = len(folder_meta)
        stats.finished = time.time()
        emit(f"cancelled after folders={stats.folders_done:,} files={len(file_rows):,}")
        push_stats()
        return [], stats

    totals = _bubble_folder_bytes(local_bytes, crawl_root_unc)

    rows: list[IndexRow] = list(file_rows)
    for unc, created in folder_meta.items():
        rows.append(
            IndexRow(
                path=unc,
                file_date=created,
                size_mb=bytes_to_size_mb(totals.get(unc, 0)),
                entry_type=ENTRY_FOLDER,
                mtime_ts=folder_mtime.get(unc, 0.0),
                size_bytes=int(local_bytes.get(unc, 0)),
            )
        )

    retried = _retry_missing_file_sizes(rows, on_progress=emit)
    if retried:
        emit(f"size retry: updated {retried:,} file(s) that had missing/zero stats")

    stats.files_indexed = sum(1 for r in rows if r.entry_type == ENTRY_FILE)
    stats.folders_indexed = sum(1 for r in rows if r.entry_type == ENTRY_FOLDER)
    stats.bytes_all_files = sum(local_bytes.values())
    stats.pending = 0
    stats.finished = time.time()
    emit(
        f"done files={stats.files_indexed:,} folders={stats.folders_indexed:,} "
        f"scan={stats.folders_scanned:,} quick={stats.folders_quick:,} "
        f"errors={stats.errors:,} elapsed={stats.elapsed:.1f}s "
        f"disk_bytes={stats.bytes_all_files:,}"
    )
    push_stats()
    return rows, stats


def _retry_missing_file_sizes(
    rows: list[IndexRow],
    *,
    on_progress: ProgressCb | None = None,
    attempts: int = 3,
    delay_sec: float = 0.15,
) -> int:
    """
    Re-stat FILE rows whose first scandir.stat failed.

    Do not retry files that already have a known size. bytes_to_size_mb floors
    anything under ~10 KB to 0.01 MB, so treating that floor as "missing" would
    re-hit thousands of tiny files and never update them.
    """
    need_idx = [
        i
        for i, r in enumerate(rows)
        if r.entry_type == ENTRY_FILE and r.size_unknown
    ]
    if not need_idx:
        return 0

    def emit(msg: str) -> None:
        if on_progress:
            on_progress(msg)

    emit(f"size retry: re-statting {len(need_idx):,} file(s) whose first stat failed...")
    updated = 0
    for n, i in enumerate(need_idx, start=1):
        path = rows[i].path
        size_bytes = 0
        for attempt in range(max(1, attempts)):
            try:
                size_bytes = int(os.stat(path, follow_symlinks=False).st_size)
                break
            except OSError:
                size_bytes = 0
            if attempt + 1 < attempts:
                time.sleep(delay_sec)
        rows[i] = IndexRow(
            path=rows[i].path,
            file_date=rows[i].file_date,
            size_mb=bytes_to_size_mb(size_bytes),
            entry_type=rows[i].entry_type,
            size_unknown=False,
        )
        if size_bytes > 0:
            updated += 1
        if n % 100 == 0:
            emit(f"size retry: {n:,}/{len(need_idx):,} checked, updated={updated:,}")
    return updated


def write_csv(rows: list[IndexRow], out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["FilePath", "FileDate", "SizeMB", "EntryType"])
        for r in rows:
            w.writerow([r.path, r.file_date, f"{r.size_mb:.2f}", r.entry_type])


def accdb_file_name_for_date(snapshot_date: dt.date | None = None) -> str:
    """AccDB filename for a scrape snapshot, e.g. SearchIndex-8-13-2026.accdb (no zero-padding)."""
    d = snapshot_date or dt.date.today()
    return f"{ACCDB_FILE_PREFIX}-{d.month}-{d.day}-{d.year}.accdb"


def parse_search_index_date(file_name: str) -> dt.date | None:
    """Parse SearchIndex-M-D-YYYY.accdb (mirrors VBA ParseSearchIndexDate)."""
    stem = str(file_name).strip()
    if stem.lower().endswith(".accdb"):
        stem = stem[:-6]
    prefix = f"{ACCDB_FILE_PREFIX}-"
    if not stem.lower().startswith(prefix.lower()):
        return None
    parts = stem[len(prefix) :].split("-")
    if len(parts) != 3:
        return None
    try:
        month = int(parts[0])
        day = int(parts[1])
        year = int(parts[2])
    except ValueError:
        return None
    if month < 1 or month > 12 or day < 1 or day > 31 or year < 1900:
        return None
    try:
        return dt.date(year, month, day)
    except ValueError:
        return None


def format_snapshot_date(d: dt.date) -> str:
    return f"{d.strftime('%B')} {d.day}, {d.year}"


def accdb_dir_for_workbook(workbook: str | Path) -> Path:
    return Path(workbook).resolve().parent / ACCDB_DIR_NAME


def default_workbook_path() -> Path:
    return Path(__file__).resolve().parent.parent.parent / DEFAULT_WORKBOOK_NAME


def _dated_search_indexes(db_dir: Path) -> list[tuple[dt.date, Path]]:
    found: list[tuple[dt.date, Path]] = []
    if not db_dir.is_dir():
        return found
    for path in db_dir.iterdir():
        if not path.is_file() or path.suffix.lower() != ".accdb":
            continue
        parsed = parse_search_index_date(path.name)
        if parsed is not None:
            found.append((parsed, path))
    return found


def existing_accdb_for_workbook(workbook: str | Path, *, warn: bool = True) -> Path | None:
    """Newest dated SearchIndex-* else LAN_Search_Index.accdb. None if missing (fresh create)."""
    db_dir = accdb_dir_for_workbook(workbook)
    dated = _dated_search_indexes(db_dir)
    if dated:
        if warn and len(dated) > 1:
            names = ", ".join(p.name for _, p in dated)
            print(
                f"Warning: {len(dated)} dated AccDB files in {db_dir}; "
                f"using the newest. Leftovers: {names}",
                flush=True,
            )

        def _rank(item: tuple[dt.date, Path]) -> tuple[dt.date, float]:
            path = item[1]
            try:
                mtime = path.stat().st_mtime
            except OSError:
                mtime = 0.0
            return (item[0], mtime)

        return max(dated, key=_rank)[1]
    legacy = db_dir / ACCDB_FILE_NAME_LEGACY
    if legacy.is_file():
        return legacy
    return None


def format_accdb_taken_on(path: Path) -> str:
    parsed = parse_search_index_date(path.name)
    if parsed is not None:
        return format_snapshot_date(parsed)
    try:
        stamp = dt.datetime.fromtimestamp(path.stat().st_mtime).date()
    except OSError:
        stamp = dt.date.today()
    return format_snapshot_date(stamp)


def accdb_path_for_workbook(
    workbook: str | Path,
    snapshot_date: dt.date | None = None,
) -> Path:
    """Existing AccDB beside the workbook, or today's dated name if none exists."""
    existing = existing_accdb_for_workbook(workbook)
    if existing is not None:
        return existing
    return accdb_dir_for_workbook(workbook) / accdb_file_name_for_date(snapshot_date)


def purge_workbook_accdbs(workbook: str | Path) -> None:
    """Delete SearchIndex-*.accdb, LAN_Search_Index.accdb, and .laccdb locks in DB\\."""
    db_dir = accdb_dir_for_workbook(workbook)
    if not db_dir.is_dir():
        return
    targets: list[Path] = []
    for path in db_dir.iterdir():
        if not path.is_file():
            continue
        suffix = path.suffix.lower()
        if suffix == ".laccdb":
            targets.append(path)
            continue
        if suffix != ".accdb":
            continue
        if parse_search_index_date(path.name) is not None:
            targets.append(path)
        elif path.name.lower() == ACCDB_FILE_NAME_LEGACY.lower():
            targets.append(path)
    failed: list[str] = []
    for path in targets:
        try:
            path.unlink()
        except FileNotFoundError:
            continue
        except OSError as exc:
            failed.append(f"{path.name}: {exc}")
    if failed:
        raise RuntimeError(
            "Could not delete AccDB (close Lan_Search_Tool.xlsm / Excel and retry). "
            + "; ".join(failed)
        )


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


def _python_bitness() -> int:
    return struct.calcsize("P") * 8


def _ace_missing_message(last_err: Exception | None) -> str:
    bits = _python_bitness()
    other = 32 if bits == 64 else 64
    return (
        f"Could not open AccDB (Python is {bits}-bit). Install Microsoft Access "
        f"Database Engine (ACE) {bits}-bit so it matches this Python — not {other}-bit Office.\n"
        "Download: Microsoft Access Database Engine 2016 Redistributable "
        "(AccessDatabaseEngine_X64.exe for 64-bit Python).\n"
        "If 32-bit Office is already installed, run the 64-bit ACE installer from an "
        "elevated command prompt:\n"
        "  AccessDatabaseEngine_X64.exe /quiet\n"
        f"Last error: {last_err}"
    )


def _ado_append_param(cmd: Any, value: object) -> None:
    ad_param_input = 1
    ad_integer = 3
    ad_double = 5
    ad_long_var_wchar = 203
    if value is None:
        param = cmd.CreateParameter("", ad_long_var_wchar, ad_param_input, 1)
        param.Value = None
    elif isinstance(value, bool):
        param = cmd.CreateParameter("", ad_integer, ad_param_input, 0, int(value))
    elif isinstance(value, int) and not isinstance(value, bool):
        param = cmd.CreateParameter("", ad_integer, ad_param_input, 0, value)
    elif isinstance(value, float):
        param = cmd.CreateParameter("", ad_double, ad_param_input, 0, value)
    else:
        text = str(value)
        param = cmd.CreateParameter(
            "", ad_long_var_wchar, ad_param_input, max(len(text), 1), text
        )
    cmd.Parameters.Append(param)


class _AdoCursor:
    def __init__(self, conn: Any) -> None:
        self._conn = conn
        self._rs: Any = None
        self.rowcount = -1

    def execute(self, sql: str, params: tuple[object, ...] | list[object] | None = None) -> None:
        import win32com.client  # type: ignore

        cmd = win32com.client.Dispatch("ADODB.Command")
        cmd.ActiveConnection = self._conn
        cmd.CommandText = sql
        cmd.CommandType = 1  # adCmdText
        if params:
            for value in params:
                _ado_append_param(cmd, value)
        self._rs, self.rowcount = cmd.Execute()

    def executemany(self, sql: str, seq_of_params: list[tuple[object, ...]]) -> None:
        for params in seq_of_params:
            self.execute(sql, params)
        self.rowcount = len(seq_of_params)

    def fetchone(self) -> tuple[object, ...] | None:
        rs = self._rs
        if rs is None:
            return None
        try:
            if rs.EOF:
                return None
        except Exception:  # noqa: BLE001
            return None
        row = tuple(rs.Fields.Item(i).Value for i in range(rs.Fields.Count))
        rs.MoveNext()
        return row

    def fetchall(self) -> list[tuple[object, ...]]:
        rows: list[tuple[object, ...]] = []
        while True:
            row = self.fetchone()
            if row is None:
                break
            rows.append(row)
        return rows


class _AdoConnection:
    """ACE OLEDB via ADODB when Access ODBC is missing (common with 64-bit Python)."""

    def __init__(self, conn: Any) -> None:
        self._conn = conn
        try:
            conn.BeginTrans()
            self._in_trans = True
        except Exception:  # noqa: BLE001
            self._in_trans = False

    def cursor(self) -> _AdoCursor:
        return _AdoCursor(self._conn)

    def commit(self) -> None:
        if not self._in_trans:
            return
        self._conn.CommitTrans()
        self._conn.BeginTrans()

    def rollback(self) -> None:
        if not self._in_trans:
            return
        try:
            self._conn.RollbackTrans()
        except Exception:  # noqa: BLE001
            pass
        try:
            self._conn.BeginTrans()
        except Exception:  # noqa: BLE001
            self._in_trans = False

    def close(self) -> None:
        try:
            if self._in_trans:
                self._conn.CommitTrans()
        except Exception:  # noqa: BLE001
            pass
        try:
            self._conn.Close()
        except Exception:  # noqa: BLE001
            pass


def _open_accdb_ado(out_path: Path) -> _AdoConnection | None:
    try:
        import win32com.client  # type: ignore
    except ImportError:
        return None
    path = str(Path(out_path).resolve())
    for provider in ("Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0"):
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(f"Provider={provider};Data Source={path};")
            return _AdoConnection(conn)
        except Exception:  # noqa: BLE001
            continue
    return None


@dataclass
class IngestLogRow:
    """Mirrors Ingestion!tblIngested (Links already Eaten)."""

    root_path: str
    date_ingested: str  # yyyy-mm-dd
    file_count: int
    folder_count: int
    total_size_mb: float


def ingest_stats_from_rows(rows: list[IndexRow]) -> tuple[int, int, float]:
    """FileCount, FolderCount, TotalSizeMB — FILE SizeMB sum (matches VBA / sheet import)."""
    files = 0
    folders = 0
    size_mb = 0.0
    for r in rows:
        if r.entry_type == ENTRY_FOLDER:
            folders += 1
        else:
            files += 1
            size_mb += float(r.size_mb)
    return files, folders, round(size_mb, 2)


def _norm_ingest_root(path: str) -> str:
    return strip_trailing_slash(path).lower()


def _is_synthetic_ingest_root(path: str) -> bool:
    return strip_trailing_slash(path).lower() in ("(accdb index)", "")


def _merge_ingest_log(
    existing: list[IngestLogRow],
    new_row: IngestLogRow,
) -> list[IngestLogRow]:
    """Upsert one scan root; drop exact match + child roots (ReplaceRoot)."""
    parent = strip_trailing_slash(new_row.root_path)
    parent_key = _norm_ingest_root(parent)
    keep: list[IngestLogRow] = []
    for row in existing:
        path = strip_trailing_slash(row.root_path)
        if _is_synthetic_ingest_root(path):
            continue
        key = _norm_ingest_root(path)
        if key == parent_key:
            continue
        if path_starts_with_root(path, parent):
            continue
        keep.append(row)
    keep.append(
        IngestLogRow(
            root_path=parent,
            date_ingested=new_row.date_ingested,
            file_count=new_row.file_count,
            folder_count=new_row.folder_count,
            total_size_mb=new_row.total_size_mb,
        )
    )
    keep.sort(key=lambda r: r.root_path.lower())
    return keep


def _accdb_table_exists(cur, table_name: str) -> bool:
    try:
        cur.execute(f"SELECT TOP 1 * FROM [{table_name}]")
        cur.fetchone()
        return True
    except Exception:  # noqa: BLE001
        return False


def _ensure_accdb_tables(cur, conn) -> None:
    if not _accdb_table_exists(cur, "tblFiles"):
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
        try:
            cur.execute("CREATE INDEX ix_tblFiles_path ON tblFiles (FilePath)")
            cur.execute("CREATE INDEX ix_tblFiles_type ON tblFiles (EntryType)")
            conn.commit()
        except Exception:  # noqa: BLE001
            conn.rollback()

    if not _accdb_table_exists(cur, "tblFolderMeta"):
        cur.execute(
            """
            CREATE TABLE tblFolderMeta (
                FolderPath TEXT NOT NULL,
                DirMtime DOUBLE,
                LocalBytes DOUBLE,
                CreatedDate TEXT
            )
            """
        )
        conn.commit()
        try:
            cur.execute("CREATE INDEX ix_tblFolderMeta_path ON tblFolderMeta (FolderPath)")
            conn.commit()
        except Exception:  # noqa: BLE001
            conn.rollback()

    if not _accdb_table_exists(cur, "tblIngested"):
        cur.execute(
            """
            CREATE TABLE tblIngested (
                RootPath TEXT NOT NULL,
                DateIngested TEXT,
                FileCount INTEGER,
                FolderCount INTEGER,
                TotalSizeMB DOUBLE
            )
            """
        )
        conn.commit()
    """Open AccDB via pyodbc, else ACE OLEDB. Caller must close."""
    try:
        import pyodbc  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "Writing AccDB requires pyodbc. Install: pip install pyodbc"
        ) from exc

    last_err: Exception | None = None
    for conn_str in _accdb_connection_strings(out_path):
        try:
            return pyodbc.connect(conn_str, autocommit=False)
        except Exception as exc:  # noqa: BLE001
            last_err = exc
    ado = _open_accdb_ado(out_path)
    if ado is not None:
        return ado
    raise RuntimeError(_ace_missing_message(last_err)) from last_err


def _under_root_params(root: str) -> tuple[str, int, str]:
    root = strip_trailing_slash(root)
    prefix = root + "\\"
    return root.upper(), len(prefix), prefix.upper()


def load_previous_index(accdb_path: str | Path, crawl_root: str) -> PreviousIndex | None:
    """Load tblFiles + tblFolderMeta under crawl_root. None if the AccDB is missing."""
    path = Path(accdb_path)
    if not path.is_file():
        return None
    root = strip_trailing_slash(crawl_root)
    if not root:
        return None

    conn = _open_accdb(path)
    try:
        cur = conn.cursor()
        if not _accdb_table_exists(cur, "tblFiles"):
            return PreviousIndex(has_meta=False)

        root_u, prefix_len, prefix_u = _under_root_params(root)
        cur.execute(
            "SELECT FilePath, FileDate, SizeMB, EntryType FROM tblFiles "
            "WHERE UCase(FilePath)=? OR Left(UCase(FilePath),?)=?",
            (root_u, prefix_len, prefix_u),
        )
        file_rows: list[tuple[str, str, float, str]] = []
        for rec in cur.fetchall():
            path_txt = strip_trailing_slash(str(rec[0] or "").strip())
            date_val = rec[1]
            if date_val is None:
                date_txt = ""
            elif hasattr(date_val, "isoformat"):
                date_txt = date_val.isoformat()[:10]
            else:
                date_txt = str(date_val).strip()[:10]
            try:
                size_mb = float(rec[2] or 0)
            except (TypeError, ValueError):
                size_mb = 0.0
            et = str(rec[3] or "FILE").strip()
            file_rows.append((path_txt, date_txt, size_mb, et))

        meta_rows: list[FolderMetaRow] = []
        if _accdb_table_exists(cur, "tblFolderMeta"):
            cur.execute(
                "SELECT FolderPath, DirMtime, LocalBytes, CreatedDate FROM tblFolderMeta "
                "WHERE UCase(FolderPath)=? OR Left(UCase(FolderPath),?)=?",
                (root_u, prefix_len, prefix_u),
            )
            for rec in cur.fetchall():
                folder_path = strip_trailing_slash(str(rec[0] or "").strip())
                if not folder_path:
                    continue
                try:
                    dir_mtime = float(rec[1] or 0)
                except (TypeError, ValueError):
                    dir_mtime = 0.0
                try:
                    local_b = int(float(rec[2] or 0))
                except (TypeError, ValueError):
                    local_b = 0
                created = ""
                if rec[3] is not None:
                    created = str(rec[3]).strip()[:10]
                meta_rows.append(
                    FolderMetaRow(
                        folder_path=folder_path,
                        dir_mtime=dir_mtime,
                        local_bytes=local_b,
                        created_date=created,
                    )
                )

        return build_previous_index(file_rows=file_rows, meta_rows=meta_rows)
    finally:
        conn.close()


def _delete_folder_meta_under_root(cur, root: str) -> None:
    if not _accdb_table_exists(cur, "tblFolderMeta"):
        return
    root = strip_trailing_slash(root)
    if not root:
        return
    root_u, prefix_len, prefix_u = _under_root_params(root)
    cur.execute(
        "DELETE FROM tblFolderMeta WHERE UCase(FolderPath)=? OR Left(UCase(FolderPath),?)=?",
        (root_u, prefix_len, prefix_u),
    )


def _insert_folder_meta_rows(cur, conn, rows: list[IndexRow]) -> None:
    folders = [r for r in rows if r.entry_type == ENTRY_FOLDER]
    if not folders:
        return
    insert_sql = (
        "INSERT INTO tblFolderMeta (FolderPath, DirMtime, LocalBytes, CreatedDate) "
        "VALUES (?, ?, ?, ?)"
    )
    batch: list[tuple[str, float, float, str | None]] = []
    batch_size = 2000
    for r in folders:
        batch.append(
            (
                strip_trailing_slash(r.path),
                float(r.mtime_ts or 0.0),
                float(r.size_bytes or 0),
                r.file_date or None,
            )
        )
        if len(batch) >= batch_size:
            cur.executemany(insert_sql, batch)
            conn.commit()
            batch.clear()
    if batch:
        cur.executemany(insert_sql, batch)
        conn.commit()


def _read_accdb_ingest_log(cur) -> list[IngestLogRow]:
    if not _accdb_table_exists(cur, "tblIngested"):
        return []
    rows: list[IngestLogRow] = []
    cur.execute(
        "SELECT RootPath, DateIngested, FileCount, FolderCount, TotalSizeMB "
        "FROM tblIngested ORDER BY RootPath"
    )
    for rec in cur.fetchall():
        root = strip_trailing_slash(str(rec[0] or "").strip())
        if _is_synthetic_ingest_root(root):
            continue
        date_val = rec[1]
        if date_val is None:
            date_txt = ""
        elif hasattr(date_val, "isoformat"):
            date_txt = date_val.isoformat()[:10]
        else:
            date_txt = str(date_val).strip()[:10]
        try:
            files = int(rec[2] or 0)
        except (TypeError, ValueError):
            files = 0
        try:
            folders = int(rec[3] or 0)
        except (TypeError, ValueError):
            folders = 0
        try:
            size_mb = round(float(rec[4] or 0), 2)
        except (TypeError, ValueError):
            size_mb = 0.0
        rows.append(
            IngestLogRow(
                root_path=root,
                date_ingested=date_txt,
                file_count=files,
                folder_count=folders,
                total_size_mb=size_mb,
            )
        )
    return rows


def _delete_accdb_files_under_root(cur, root: str) -> int:
    """Delete tblFiles rows for this root (exact + descendants). Returns deleted estimate."""
    root = strip_trailing_slash(root)
    if not root:
        return 0
    prefix = root + "\\"
    root_u = root.upper()
    prefix_u = prefix.upper()
    cur.execute(
        "DELETE FROM tblFiles WHERE UCase(FilePath)=? OR Left(UCase(FilePath),?)=?",
        (root_u, len(prefix_u), prefix_u),
    )
    return cur.rowcount if cur.rowcount is not None and cur.rowcount >= 0 else 0


def _rewrite_accdb_ingest_log(cur, conn, log_rows: list[IngestLogRow]) -> None:
    try:
        cur.execute("DELETE FROM tblIngested")
        conn.commit()
    except Exception:  # noqa: BLE001
        conn.rollback()
        cur.execute("DROP TABLE tblIngested")
        conn.commit()
        cur.execute(
            """
            CREATE TABLE tblIngested (
                RootPath TEXT NOT NULL,
                DateIngested TEXT,
                FileCount INTEGER,
                FolderCount INTEGER,
                TotalSizeMB DOUBLE
            )
            """
        )
        conn.commit()

    if not log_rows:
        return
    cur.executemany(
        "INSERT INTO tblIngested "
        "(RootPath, DateIngested, FileCount, FolderCount, TotalSizeMB) "
        "VALUES (?, ?, ?, ?, ?)",
        [
            (
                strip_trailing_slash(r.root_path),
                r.date_ingested or None,
                int(r.file_count),
                int(r.folder_count),
                float(r.total_size_mb),
            )
            for r in log_rows
            if not _is_synthetic_ingest_root(r.root_path)
        ],
    )
    conn.commit()


def _insert_accdb_file_rows(cur, conn, rows: list[IndexRow]) -> None:
    insert_sql = (
        "INSERT INTO tblFiles (FilePath, FileDate, SizeMB, EntryType) VALUES (?, ?, ?, ?)"
    )
    batch: list[tuple[str, str | None, float, str]] = []
    batch_size = 2000
    for r in rows:
        batch.append((r.path, r.file_date or None, float(r.size_mb), r.entry_type))
        if len(batch) >= batch_size:
            cur.executemany(insert_sql, batch)
            conn.commit()
            batch.clear()
    if batch:
        cur.executemany(insert_sql, batch)
        conn.commit()


def write_accdb(
    rows: list[IndexRow],
    out_path: Path,
    *,
    crawl_root: str | None = None,
    ingest_log: list[IngestLogRow] | None = None,
    replace_root: bool = True,
) -> None:
    """
    Write crawl rows into AccDB tblFiles and accumulate scan roots in tblIngested.

    Default replace_root=True: keep other roots' files + ingest history; replace only
    this crawl_root (and child ingest entries). Full wipe only when replace_root=False
    or the AccDB file is new.
    """
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    root = strip_trailing_slash((crawl_root or "").strip())
    files, folders, size_mb = ingest_stats_from_rows(rows)
    new_ingest = IngestLogRow(
        root_path=root,
        date_ingested=dt.date.today().isoformat(),
        file_count=files,
        folder_count=folders,
        total_size_mb=size_mb,
    ) if root else None

    created_new = False
    if not out_path.exists():
        _create_empty_accdb(out_path)
        created_new = True

    conn = _open_accdb(out_path)
    try:
        cur = conn.cursor()
        _ensure_accdb_tables(cur, conn)

        existing_log = _read_accdb_ingest_log(cur) if not created_new else []

        full_rebuild = created_new or (not replace_root) or (not root)
        if full_rebuild:
            try:
                cur.execute("DELETE FROM tblFiles")
                conn.commit()
            except Exception:  # noqa: BLE001
                conn.rollback()
            try:
                cur.execute("DELETE FROM tblFolderMeta")
                conn.commit()
            except Exception:  # noqa: BLE001
                conn.rollback()
            _insert_accdb_file_rows(cur, conn, rows)
            _insert_folder_meta_rows(cur, conn, rows)
            if ingest_log is not None:
                log_rows = [
                    r for r in ingest_log if not _is_synthetic_ingest_root(r.root_path)
                ]
            elif new_ingest is not None:
                log_rows = [new_ingest]
            else:
                log_rows = []
            _rewrite_accdb_ingest_log(cur, conn, log_rows)
            print(
                f"AccDB full write: files={len(rows):,} ingest_scans={len(log_rows)}",
                flush=True,
            )
            return

        # ReplaceRoot merge: keep other roots, replace this root's files + ingest row
        deleted = _delete_accdb_files_under_root(cur, root)
        _delete_folder_meta_under_root(cur, root)
        conn.commit()
        _insert_accdb_file_rows(cur, conn, rows)
        _insert_folder_meta_rows(cur, conn, rows)
        assert new_ingest is not None
        if ingest_log is not None:
            log_rows = [
                r for r in ingest_log if not _is_synthetic_ingest_root(r.root_path)
            ]
        else:
            log_rows = _merge_ingest_log(existing_log, new_ingest)
        _rewrite_accdb_ingest_log(cur, conn, log_rows)
        print(
            f"AccDB ReplaceRoot: removed~{deleted} prior rows under {root}; "
            f"ingest scans now={len(log_rows)}",
            flush=True,
        )
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


def remove_crawl_output_dir(out_dir: Path | None = None) -> None:
    """Remove leftover crawl_output folders (CSV is no longer staged for AccDB)."""
    path = Path(out_dir) if out_dir is not None else _default_out_dir()
    if not path.is_dir():
        return
    try:
        shutil.rmtree(path)
        print(f"Removed temporary {path}", flush=True)
    except OSError as exc:
        print(f"Could not remove {path}: {exc}", flush=True)


def run_interactive(
    *,
    workers: int = 16,
    progress_every: int = 50,
    auto_import_excel: bool = True,
    workbook: str | Path | None = None,
    target_mode: str | None = None,
    use_gui: bool = True,
    accdb_fresh: bool = False,
    incremental: bool = True,
    full_crawl: bool = False,
) -> int:
    """GUI/CLI entry: crawl → AccDB ReplaceRoot merge (or fresh snapshot)."""
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
    except Exception:
        pass

    def say(msg: str) -> None:
        if not live_gui:
            print(msg, flush=True)

    live_gui = False
    if not use_gui:
        say("LAN Search Tool - parallel crawler")

    root = ""
    wb_path: Path | None = None
    mode = (target_mode or "AccDB").strip() or "AccDB"
    if mode not in ("Auto", "Workbook", "AccDB"):
        mode = "AccDB"
    want_fresh = bool(accdb_fresh)
    want_incremental = bool(incremental) and not full_crawl and not want_fresh
    workers = max(1, min(48, int(workers)))

    if use_gui and workbook is None and target_mode is None:
        try:
            from .crawl_gui import run_crawl_gui, run_progress_window
        except ImportError:
            from crawler.crawl_gui import run_crawl_gui, run_progress_window  # type: ignore

        gui = run_crawl_gui(initial_workbook=workbook)
        if gui.cancelled:
            return 0
        root = gui.root
        wb_path = Path(gui.workbook)
        mode = gui.target_mode
        workers = gui.workers
        want_fresh = gui.accdb_fresh
        want_incremental = gui.incremental and not want_fresh
        live_gui = True
    else:
        print("Opening folder picker (check the taskbar if you do not see it)...", flush=True)
        root = pick_folder()
        if not root:
            print("Cancelled - no folder selected.", flush=True)
            return 0

        if workbook is not None and str(workbook).strip():
            wb_path = Path(workbook)
        elif auto_import_excel:
            default_wb = default_workbook_path()
            if default_wb.is_file():
                wb_path = default_wb
            else:
                print("Opening workbook picker (select the .xlsm to update)...", flush=True)
                chosen = pick_workbook(
                    initial_dir=Path(__file__).resolve().parent.parent.parent,
                    title="Select LAN Search workbook",
                )
                if not chosen:
                    print("Cancelled - no workbook selected.", flush=True)
                    return 0
                wb_path = Path(chosen)
        else:
            print("No workbook selected; AccDB/Excel write skipped.", flush=True)
            wb_path = None

    if wb_path is not None and not wb_path.is_file():
        msg = f"Workbook not found: {wb_path}"
        if live_gui:
            from crawler.crawl_gui import report_fatal

            report_fatal("LAN Search Tool", msg)
        else:
            print(msg, flush=True)
        return 1

    if want_fresh and wb_path is not None:
        say("Starting fresh: deleting existing AccDB files in DB\\ ...")
        try:
            purge_workbook_accdbs(wb_path)
        except RuntimeError as exc:
            if live_gui:
                from crawler.crawl_gui import report_fatal

                report_fatal("LAN Search Tool", str(exc))
            else:
                print(str(exc), flush=True)
            return 1

    accdb_out = accdb_path_for_workbook(wb_path) if wb_path is not None else None

    drive_map = build_drive_map([root])
    unc = to_unc_path(root, drive_map, "")
    crawl_mode_label = (
        "Start fresh (wipe index)"
        if want_fresh
        else ("Quick update" if want_incremental else "Full recrawl")
    )
    say(f"Root:    {root}")
    say(f"UNC:     {unc}")
    say(f"Workers: {workers}")
    say(f"Crawl:   {crawl_mode_label}")
    say(f"Mode:    {mode}")
    if wb_path is not None:
        say(f"Excel:   {wb_path}")
    if accdb_out is not None:
        say(f"AccDB:   {accdb_out}")
    say("---")

    job: dict[str, object] = {"rows": None, "stats": None, "error": None}

    def run_job(
        cancel_event: threading.Event | None,
        on_stats: StatsCb | None,
        on_log: ProgressCb | None,
    ) -> None:
        def emit(msg: str) -> None:
            if not live_gui:
                print(msg, flush=True)
            if on_log:
                on_log(msg)

        previous = None
        use_inc = want_incremental
        if use_inc and accdb_out is not None and accdb_out.is_file():
            emit("Loading previous index for quick update...")
            try:
                previous = load_previous_index(accdb_out, unc)
            except Exception as exc:  # noqa: BLE001
                emit(f"Could not load previous index ({exc}); doing a full listing.")
                previous = None
            if previous is None or not previous.has_meta:
                emit(
                    "No folder timestamps in this AccDB yet — full listing this run. "
                    "Next quick update will skip unchanged folders."
                )
                use_inc = bool(previous is not None and previous.has_meta)
            elif previous.has_meta:
                emit(
                    f"Quick update ready: {len(previous.folders):,} folders with timestamps."
                )
        elif use_inc:
            use_inc = False

        t0 = time.time()
        rows, stats = crawl_parallel(
            root,
            workers=workers,
            drive_map=drive_map,
            progress_every=progress_every,
            on_progress=emit,
            on_stats=on_stats,
            incremental=use_inc,
            previous=previous,
            cancel_event=cancel_event,
        )
        job["rows"] = rows
        job["stats"] = stats
        if stats.cancelled or (cancel_event is not None and cancel_event.is_set()):
            emit("Cancelled. Index was not updated.")
            return
        emit(
            f"Crawl done: {len(rows):,} rows  "
            f"files={stats.files_indexed:,} folders={stats.folders_indexed:,} "
            f"scan={stats.folders_scanned:,} quick={stats.folders_quick:,} "
            f"errors={stats.errors:,} disk~{bytes_to_size_mb(stats.bytes_all_files):,.2f} MB "
            f"in {time.time() - t0:.1f}s"
        )

        if wb_path is None:
            out_dir = _default_out_dir()
            stamp = time.strftime("%Y%m%d_%H%M%S")
            csv_path = out_dir / f"{stamp}_tblFiles.csv"
            write_csv(rows, csv_path)
            emit(f"Wrote CSV {csv_path} ({len(rows):,} rows).")
            return

        force_accdb = mode == "AccDB" or len(rows) > ACCDB_ROW_THRESHOLD
        if mode == "Workbook" and len(rows) > ACCDB_ROW_THRESHOLD:
            force_accdb = True
            emit(
                f"Row count {len(rows):,} exceeds {ACCDB_ROW_THRESHOLD:,}; "
                "forcing AccDB write for performance."
            )
        elif mode == "Auto" and len(rows) > ACCDB_ROW_THRESHOLD:
            force_accdb = True
            emit(
                f"Auto mode: {len(rows):,} rows > {ACCDB_ROW_THRESHOLD:,} → AccDB."
            )

        if force_accdb:
            assert accdb_out is not None
            emit(f"Writing AccDB -> {accdb_out} ...")
            write_accdb(rows, accdb_out, crawl_root=unc)
            emit(f"Wrote AccDB ({len(rows):,} rows + folder timestamps).")
            try:
                from .import_to_excel import clear_onboard_tblfiles
            except ImportError:
                from crawler.import_to_excel import clear_onboard_tblfiles  # type: ignore

            emit("Clearing onboard Database!tblFiles so AccDB is exclusive...")
            clear_onboard_tblfiles(wb_path)
            emit(
                f"Index deployed to AccDB:\n  {accdb_out}\n"
                "Re-open the workbook to search."
            )
            return

        try:
            from .import_to_excel import import_csv_to_workbook
        except ImportError:
            from crawler.import_to_excel import import_csv_to_workbook  # type: ignore

        tmp = tempfile.NamedTemporaryFile(
            prefix="lan_search_",
            suffix=".csv",
            delete=False,
        )
        csv_path = Path(tmp.name)
        tmp.close()
        try:
            write_csv(rows, csv_path)
            emit(f"Importing (ReplaceRoot) into workbook {wb_path} ...")
            total = import_csv_to_workbook(
                workbook=wb_path,
                csv_path=csv_path,
                mode="ReplaceRoot",
                unc_root=unc,
            )
        finally:
            try:
                csv_path.unlink()
            except OSError:
                pass
        emit(f"Excel import complete ({total:,} rows in tblFiles).")

    if live_gui:
        try:
            from .crawl_gui import run_progress_window
        except ImportError:
            from crawler.crawl_gui import run_progress_window  # type: ignore

        ok = run_progress_window(
            root_label=unc or root,
            mode_label=f"{crawl_mode_label}  ·  {workers} workers",
            work=run_job,
        )
        err = job.get("error")
        stats_obj = job.get("stats")
        if isinstance(stats_obj, CrawlStats) and stats_obj.cancelled:
            return 0
        if not ok:
            return 1
        if err:
            from crawler.crawl_gui import report_fatal

            report_fatal("LAN Search Tool", f"Index write failed: {err}")
            return 1
        return 0

    try:
        run_job(None, None, None)
    except Exception as exc:  # noqa: BLE001
        print(f"Index write failed: {exc}", flush=True)
        return 1
    stats_obj = job.get("stats")
    if isinstance(stats_obj, CrawlStats) and stats_obj.cancelled:
        return 0
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(run_interactive())
    except SystemExit:
        raise
    except Exception as exc:  # noqa: BLE001
        import traceback

        try:
            from crawler.crawl_gui import report_fatal

            report_fatal("LAN Search Tool", f"{exc}")
        except Exception:
            print("CRASH:", exc, flush=True)
            traceback.print_exc()
        raise SystemExit(1) from exc
