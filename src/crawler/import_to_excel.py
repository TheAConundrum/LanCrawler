"""Import crawler CSV into Lan_Search_Tool.xlsm (Database + Ingestion log)."""

from __future__ import annotations

import csv
import datetime as dt
from pathlib import Path

from workbook_security import (
    enforce_workbook_ui_lock,
    reopen_workbook_in_app,
)

from .paths import bytes_to_size_mb, path_starts_with_root, strip_trailing_slash

INGESTION_SHEET = "Ingestion"
INGESTED_TABLE = "tblIngested"
INGESTED_HEADER_ROW = 6
INGESTED_COLS = 5
SUMMARY_SCANS = "E2"
SUMMARY_FILES = "E3"
SUMMARY_FOLDERS = "E4"
SUMMARY_SIZE = "E5"


def _read_csv_rows(csv_path: Path) -> list[list[object]]:
    rows: list[list[object]] = []
    with csv_path.open("r", encoding="utf-8", newline="") as fh:
        reader = csv.DictReader(fh)
        for rec in reader:
            path = (rec.get("FilePath") or "").strip()
            if not path:
                continue
            date = (rec.get("FileDate") or "").strip()
            try:
                size = float(rec.get("SizeMB") or 0.01)
            except ValueError:
                size = 0.01
            size = (
                bytes_to_size_mb(size * 1048576.0)
                if size > 1000
                else max(0.01, round(size, 2))
            )
            et = (rec.get("EntryType") or "FILE").strip().upper() or "FILE"
            rows.append([path, date if date else None, size, et])
    return rows


def _as_float(value: object, default: float = 0.0) -> float:
    if value is None or value == "":
        return default
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            return float(value)
        except ValueError:
            return default
    try:
        return float(str(value))
    except (TypeError, ValueError):
        return default


def _csv_ingest_stats(rows: list[list[object]]) -> tuple[int, int, float]:
    """FileCount, FolderCount, TotalSizeMB (indexed files only — matches VBA mTotalBytes)."""
    files = 0
    folders = 0
    size_mb = 0.0
    for row in rows:
        et = str(row[3] or "FILE").upper()
        if et == "FOLDER":
            folders += 1
        else:
            files += 1
            size_mb += _as_float(row[2])
    return files, folders, round(size_mb, 2)


def _sheet_used_last_row(ws, col: int = 1) -> int:
    last = ws.cells.last_cell.row
    rng = ws.range((1, col), (last, col))
    vals = rng.options(ndim=1).value
    if vals is None:
        return 1
    if not isinstance(vals, list):
        vals = [vals]
    for i in range(len(vals), 0, -1):
        if vals[i - 1] not in (None, ""):
            return i
    return 1


def _find_table(ws, name: str):
    for tbl in list(ws.tables):
        if tbl.name == name:
            return tbl
    # Fallback: header RootPath
    for tbl in list(ws.tables):
        try:
            hdr = tbl.header_row_range.value
            first = hdr[0] if isinstance(hdr, list) else hdr
            if str(first or "").strip().lower() == "rootpath":
                return tbl
        except Exception:
            continue

    return None


def _ensure_ingested_table(ws, table_name: str = INGESTED_TABLE):
    """Return tblIngested, creating headers/table at A6 if needed."""
    tbl = _find_table(ws, table_name)
    if tbl is not None:
        return tbl

    header_row = INGESTED_HEADER_ROW
    ws.range((header_row, 1)).value = [
        ["RootPath", "DateIngested", "FileCount", "FolderCount", "TotalSizeMB"]
    ]
    # Create with one blank data row then delete body if empty — keep header+1 for ListObject
    end = header_row + 1
    src = ws.range((header_row, 1), (end, INGESTED_COLS))
    tbl = ws.tables.add(source=src, name=table_name)
    # Clear the placeholder data row if present
    try:
        if tbl.data_body_range is not None:
            tbl.data_body_range.clear_contents()
    except Exception:
        pass
    return tbl


def _read_ingested_rows(ws, tbl) -> list[list[object]]:
    rows: list[list[object]] = []
    if tbl is None:
        return rows
    body = tbl.data_body_range
    if body is None:
        return rows
    data = body.options(ndim=2).value
    if not data:
        return rows
    if not isinstance(data[0], list):
        data = [data]
    for row in data:
        path = str(row[0] or "").strip()
        if not path:
            continue
        rows.append([
            path,
            row[1] if len(row) > 1 else None,
            row[2] if len(row) > 2 else 0,
            row[3] if len(row) > 3 else 0,
            row[4] if len(row) > 4 else 0.01,
        ])
    return rows


def _norm_root(path: str) -> str:
    return strip_trailing_slash(path).lower()


def _upsert_ingested_log(
    ws,
    *,
    unc_root: str,
    file_count: int,
    folder_count: int,
    total_mb: float,
    mode: str,
    table_name: str = INGESTED_TABLE,
) -> None:
    """
    Update Links already Eaten (tblIngested) + E2:E5 summary.
    Mirrors VBA UpsertIngestedList / RemoveChildIngestedRoots / RefreshIngestedSummary.
    """
    store_root = strip_trailing_slash(unc_root)
    if not store_root and mode != "RebuildAll":
        print("Ingestion log skipped: no UNC root provided.", flush=True)
        return

    tbl = _ensure_ingested_table(ws, table_name)
    existing = _read_ingested_rows(ws, tbl)
    today = dt.date.today()

    if mode == "RebuildAll":
        keep: list[list[object]] = []
    else:
        keep = []
        parent = _norm_root(store_root)
        for row in existing:
            path = strip_trailing_slash(str(row[0] or ""))
            key = _norm_root(path)
            # Drop exact match (will re-add) and children under this root
            if key == parent:
                continue
            if path_starts_with_root(path, store_root):
                continue
            keep.append([path, row[1], row[2], row[3], row[4]])

    if store_root:
        keep.append([store_root, today, file_count, folder_count, total_mb])

    # Sort by root path (text)
    keep.sort(key=lambda r: str(r[0]).lower())

    # Rewrite table: delete ListObject, write range, recreate
    header_row = INGESTED_HEADER_ROW
    try:
        if tbl is not None and tbl.header_row_range is not None:
            header_row = tbl.header_row_range.row
    except Exception:
        pass

    for t in list(ws.tables):
        try:
            if t.name == table_name:
                t.api.Delete()
                continue
            hv = t.header_row_range.value
            first = hv[0] if isinstance(hv, list) else hv
            if str(first or "").strip().lower() == "rootpath":
                t.api.Delete()
        except Exception:
            continue

    # Clear old body area under header
    last = _sheet_used_last_row(ws, 1)
    if last > header_row:
        ws.range(
            (header_row + 1, 1), (max(last, header_row + 1), INGESTED_COLS)
        ).clear_contents()

    ws.range((header_row, 1)).value = [
        ["RootPath", "DateIngested", "FileCount", "FolderCount", "TotalSizeMB"]
    ]

    n = len(keep)
    if n:
        ws.range((header_row + 1, 1)).value = keep
        end_row = header_row + n
    else:
        end_row = header_row + 1  # placeholder row for empty table

    src = ws.range((header_row, 1), (end_row, INGESTED_COLS))
    new_tbl = ws.tables.add(source=src, name=table_name)
    if n == 0:
        try:
            if new_tbl.data_body_range is not None:
                new_tbl.data_body_range.clear_contents()
        except Exception:
            pass
    else:
        try:
            if new_tbl.data_body_range is not None:
                # FileCount / FolderCount / TotalSizeMB formats (optional; user may own formats)
                new_tbl.data_body_range.columns[2].number_format = "#,##0"
                new_tbl.data_body_range.columns[3].number_format = "#,##0"
                new_tbl.data_body_range.columns[4].number_format = "#,##0.00"
        except Exception:
            pass

    # Summary strip E2:E5 (values only)
    scans = n
    files = sum(_as_float(r[2]) for r in keep)
    folders = sum(_as_float(r[3]) for r in keep)
    size = round(sum(_as_float(r[4]) for r in keep), 2)
    ws.range(SUMMARY_SCANS).value = scans
    ws.range(SUMMARY_FILES).value = files
    ws.range(SUMMARY_FOLDERS).value = folders
    ws.range(SUMMARY_SIZE).value = size
    print(
        f"Ingestion log updated: root={store_root or '(rebuild)'} "
        f"files={file_count:,} folders={folder_count:,} size_mb={total_mb:,.2f} "
        f"scans_total={scans}",
        flush=True,
    )


def _workbook_keys(path: Path) -> tuple[str, str]:
    resolved = str(path.resolve()).lower()
    return resolved, path.name.lower()


def _close_target_if_open(workbook: Path):
    """
    Save+close the target workbook if open. Returns an existing Excel app to reuse,
    or None if Excel is not running.
    """
    import xlwings as xw

    target_full, target_name = _workbook_keys(workbook)
    host_app = None

    for existing_app in list(xw.apps):
        if host_app is None:
            host_app = existing_app
        for book in list(existing_app.books):
            try:
                full = str(Path(book.fullname).resolve()).lower()
            except Exception:
                full = ""
            try:
                name = book.name.lower()
            except Exception:
                name = ""
            if full == target_full or name == target_name:
                print(f"Closing open workbook for import: {book.name}...", flush=True)
                try:
                    book.save()
                except Exception:
                    pass
                try:
                    book.close()
                except Exception:
                    pass
                host_app = existing_app

    return host_app


def import_csv_to_workbook(
    workbook: Path,
    csv_path: Path,
    *,
    mode: str = "ReplaceRoot",
    unc_root: str = "",
    sheet_name: str = "Database",
    table_name: str = "tblFiles",
    ingestion_sheet: str = INGESTION_SHEET,
) -> int:
    """
    Load CSV into Database!tblFiles and update Ingestion!tblIngested + E2:E5 summary.
    mode:
      ReplaceRoot — keep existing DB rows outside unc_root; upsert this root on Ingestion
      Append — keep all existing DB rows; upsert this root on Ingestion
      RebuildAll — wipe tblFiles to CSV only; rebuild Ingestion log to this root only
    Returns total data rows written to tblFiles.

    Workflow (aligned with inject_vba silence protocol):
      close target if open → open with events/screen off → write → save → close →
      reopen with events on so Workbook_Open schedules cache warm.
    Excel stays visible but frozen during the write; left open after reopen.
    """
    import xlwings as xw

    workbook = Path(workbook)
    csv_path = Path(csv_path)
    if not workbook.exists():
        raise FileNotFoundError(workbook)
    if not csv_path.exists():
        raise FileNotFoundError(csv_path)

    new_rows = _read_csv_rows(csv_path)
    unc_root = strip_trailing_slash(unc_root)
    mode_u = mode.strip()
    file_count, folder_count, total_mb = _csv_ingest_stats(new_rows)
    saved_path = str(workbook.resolve())

    host_app = _close_target_if_open(workbook)
    app = host_app
    we_started_app = False
    if app is None:
        print(f"Opening Excel for import: {workbook.name}", flush=True)
        app = xw.App(visible=True, add_book=False)
        we_started_app = True
    else:
        print(f"Importing into Excel (events off): {workbook.name}", flush=True)

    wb = None
    try:
        # --- SILENCE PROTOCOL (no flicker / no Workbook_Open mid-write) ---
        app.visible = True
        app.screen_updating = False
        app.display_alerts = False
        try:
            app.enable_events = False
        except Exception:
            try:
                app.api.EnableEvents = False
            except Exception:
                pass
        # ------------------------------------------------------------------

        wb = app.books.open(saved_path)

        try:
            ws = wb.sheets[sheet_name]
        except Exception as exc:
            raise RuntimeError(f"Sheet '{sheet_name}' not found") from exc

        # Database may be xlSheetVeryHidden + protected from VBA UI lock
        try:
            ws.api.Visible = -1  # xlSheetVisible
        except Exception:
            pass
        try:
            ws.api.Unprotect(Password="")
        except Exception:
            try:
                ws.api.Unprotect()
            except Exception:
                pass

        # --- Database!tblFiles ---
        keep: list[list[object]] = []
        last_row = _sheet_used_last_row(ws, 1)
        if last_row >= 2 and mode_u != "RebuildAll":
            existing = ws.range((2, 1), (last_row, 4)).options(ndim=2).value
            if existing:
                if not isinstance(existing[0], list):
                    existing = [existing]
                for row in existing:
                    path = str(row[0] or "").strip()
                    if not path:
                        continue
                    if mode_u == "ReplaceRoot" and unc_root:
                        if (
                            path_starts_with_root(path, unc_root)
                            or path.lower() == unc_root.lower()
                        ):
                            continue
                    keep.append([
                        path,
                        row[1] if len(row) > 1 else None,
                        row[2] if len(row) > 2 else 0.01,
                        (str(row[3]).upper() if len(row) > 3 and row[3] else "FILE"),
                    ])

        combined = keep + new_rows

        for tbl in list(ws.tables):
            if tbl.name == table_name:
                tbl.api.Delete()

        if last_row >= 2:
            ws.range((2, 1), (last_row, 4)).clear_contents()

        ws.range("A1").value = [["FilePath", "FileDate", "SizeMB", "EntryType"]]
        total = len(combined)
        if total:
            ws.range((2, 1)).value = combined
            ws.range((2, 3), (1 + total, 3)).number_format = "#,##0.00"

        end_row = max(2, 1 + total)
        tbl_range = ws.range((1, 1), (end_row, 4))
        ws.tables.add(source=tbl_range, name=table_name)

        # --- Ingestion!tblIngested + summary ---
        try:
            ws_ing = wb.sheets[ingestion_sheet]
        except Exception as exc:
            raise RuntimeError(f"Sheet '{ingestion_sheet}' not found") from exc

        try:
            ws_ing.api.Unprotect(Password="")
        except Exception:
            try:
                ws_ing.api.Unprotect()
            except Exception:
                pass

        _upsert_ingested_log(
            ws_ing,
            unc_root=unc_root,
            file_count=file_count,
            folder_count=folder_count,
            total_mb=total_mb,
            mode=mode_u,
        )

        # Match VBA ProtectWorkbookUi + lock VBA project for viewing
        try:
            wb.sheets["Dashboard"].api.Unprotect(Password="")
        except Exception:
            try:
                wb.sheets["Dashboard"].api.Unprotect()
            except Exception:
                pass

        enforce_workbook_ui_lock(wb)

        wb.save()
        print(f"Saved {wb.name} ({total:,} tblFiles rows).", flush=True)

        print("Reopening workbook so Workbook_Open / cache warm can run...", flush=True)
        reopen_workbook_in_app(app, saved_path)
        app = None  # leave Excel running — do not Quit (avoids Safe Mode prompt)
        print("Import complete. Excel left open.", flush=True)
        return total
    finally:
        if app is not None:
            # Import failed after we started Excel — clean up
            try:
                if wb is not None:
                    wb.close()
            except Exception:
                pass
            if we_started_app:
                try:
                    from workbook_security import soft_quit_excel_app

                    soft_quit_excel_app(app)
                except Exception:
                    try:
                        app.quit()
                    except Exception:
                        pass
            else:
                try:
                    from workbook_security import restore_excel_interactive

                    restore_excel_interactive(app)
                except Exception:
                    pass


def _clear_ingested_table_for_accdb(ws) -> None:
    """Empty Links already Eaten — AccDB tblIngested will refill on cache warm."""
    tbl = _ensure_ingested_table(ws, INGESTED_TABLE)
    try:
        while tbl.api.ListRows.Count > 0:
            tbl.api.ListRows(1).Delete()
    except Exception:
        try:
            body = tbl.data_body_range
            if body is not None:
                body.clear_contents()
        except Exception:
            pass
    ws.range(SUMMARY_SCANS).value = 0
    ws.range(SUMMARY_FILES).value = 0
    ws.range(SUMMARY_FOLDERS).value = 0
    ws.range(SUMMARY_SIZE).value = 0


def clear_onboard_tblfiles(
    workbook: Path,
    *,
    sheet_name: str = "Database",
    table_name: str = "tblFiles",
) -> None:
    """
    Empty Database!tblFiles so AccDB can be the exclusive session backend.
    Also clears Ingestion!tblIngested (AccDB scan history reloads on warm).
    Required after AccDB writes so leftover sheet rows cannot shadow AccDB on load.
    """
    import xlwings as xw

    workbook = Path(workbook)
    if not workbook.exists():
        raise FileNotFoundError(workbook)

    saved_path = str(workbook.resolve())
    host_app = _close_target_if_open(workbook)
    app = host_app
    we_started_app = False
    if app is None:
        app = xw.App(visible=False, add_book=False)
        we_started_app = True

    wb = None
    try:
        app.visible = False
        app.screen_updating = False
        app.display_alerts = False
        try:
            app.enable_events = False
        except Exception:
            try:
                app.api.EnableEvents = False
            except Exception:
                pass

        wb = app.books.open(saved_path)
        try:
            ws = wb.sheets[sheet_name]
        except Exception as exc:
            raise RuntimeError(f"Sheet '{sheet_name}' not found") from exc

        try:
            ws.api.Visible = -1  # xlSheetVisible
        except Exception:
            pass
        try:
            ws.api.Unprotect(Password="")
        except Exception:
            try:
                ws.api.Unprotect()
            except Exception:
                pass

        last_row = _sheet_used_last_row(ws, 1)
        for tbl in list(ws.tables):
            if tbl.name == table_name:
                tbl.api.Delete()

        if last_row >= 2:
            ws.range((2, 1), (last_row, 4)).clear_contents()

        ws.range("A1").value = [["FilePath", "FileDate", "SizeMB", "EntryType"]]
        # Keep a valid empty ListObject (header + one blank row then delete body pattern)
        ws.range((2, 1)).value = [["", None, None, ""]]
        tbl_range = ws.range((1, 1), (2, 4))
        ws.tables.add(source=tbl_range, name=table_name)
        try:
            ws.tables[table_name].api.ListRows(1).Delete()
        except Exception:
            # If delete fails, clear the blank row contents
            ws.range((2, 1), (2, 4)).clear_contents()

        # AccDB owns scan history — wipe sheet Ingestion so AccDB tblIngested is source of truth
        try:
            ws_ing = wb.sheets[INGESTION_SHEET]
            try:
                ws_ing.api.Unprotect(Password="")
            except Exception:
                try:
                    ws_ing.api.Unprotect()
                except Exception:
                    pass
            _clear_ingested_table_for_accdb(ws_ing)
            print("Cleared Ingestion!tblIngested (will reload from AccDB on open).", flush=True)
        except Exception as exc:
            print(f"Ingestion clear skipped: {exc}", flush=True)

        try:
            wb.sheets["Dashboard"].api.Unprotect(Password="")
        except Exception:
            try:
                wb.sheets["Dashboard"].api.Unprotect()
            except Exception:
                pass

        enforce_workbook_ui_lock(wb)
        wb.save()
        print(f"Cleared onboard {table_name} in {wb.name} (AccDB is exclusive index).", flush=True)

        print("Reopening workbook so AccDB cache warm can run...", flush=True)
        reopen_workbook_in_app(app, saved_path)
        app = None  # leave Excel running — do not Quit
    finally:
        if app is not None:
            try:
                if wb is not None:
                    wb.close()
            except Exception:
                pass
            if we_started_app:
                try:
                    from workbook_security import soft_quit_excel_app

                    soft_quit_excel_app(app)
                except Exception:
                    try:
                        app.quit()
                    except Exception:
                        pass
            else:
                try:
                    from workbook_security import restore_excel_interactive

                    restore_excel_interactive(app)
                except Exception:
                    pass
