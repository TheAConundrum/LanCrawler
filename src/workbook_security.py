"""
Workbook UI lock helpers for LAN Search Tool.

Mirrors clsSheetMap.ProtectWorkbookUi for the refactored Dashboard:
  - Free-text: Contains A3:B3, Contains D3:E3, Size MB G3, File Search A5:F5
  - Dropdowns: Op C3, Size is F3, Flexible H3, Folders I3, Drive G5
  - Ingestion locked; Database locked + xlSheetVeryHidden
"""

from __future__ import annotations

import time
from typing import Any

XL_SHEET_VERY_HIDDEN = 2
XL_NO_RESTRICTIONS = 0  # xlEnableSelection.xlNoRestrictions


def _unprotect_sheet_api(api: Any) -> None:
    try:
        api.Unprotect(Password="")
    except Exception:
        try:
            api.Unprotect()
        except Exception:
            pass


def _protect_sheet_api(
    api: Any,
    *,
    allow_formatting_columns: bool = False,
    allow_filtering: bool = False,
) -> None:
    """Match clsSheetMap.ApplySheetProtect (UserInterfaceOnly for this Excel session)."""
    try:
        api.Protect(
            Password="",
            DrawingObjects=False,
            Contents=True,
            Scenarios=True,
            UserInterfaceOnly=True,
            AllowFormattingCells=False,
            AllowFormattingColumns=allow_formatting_columns,
            AllowFormattingRows=False,
            AllowInsertingColumns=False,
            AllowInsertingRows=False,
            AllowInsertingHyperlinks=False,
            AllowDeletingColumns=False,
            AllowDeletingRows=False,
            AllowSorting=allow_filtering,
            AllowFiltering=allow_filtering,
            AllowUsingPivotTables=False,
        )
    except Exception:
        try:
            api.Protect(Password="", DrawingObjects=False, Contents=True, Scenarios=True)
        except Exception as exc:
            print(f"  [!] Sheet protect failed: {exc}", flush=True)
    try:
        api.EnableSelection = XL_NO_RESTRICTIONS
    except Exception:
        pass


def _set_list_validation(api: Any, address: str, formula: str) -> None:
    """Stop-style list validation so unlocked dropdown cells reject free typing."""
    try:
        cell = api.Range(address)
        try:
            if bool(cell.MergeCells):
                cell = cell.MergeArea.Cells(1, 1)
        except Exception:
            pass
        try:
            cell.Validation.Delete()
        except Exception:
            pass
        cell.Validation.Add(
            Type=3,  # xlValidateList
            AlertStyle=1,  # xlValidAlertStop
            Operator=1,
            Formula1=formula,
        )
        cell.Validation.IgnoreBlank = True
        cell.Validation.InCellDropdown = True
        cell.Validation.ShowInput = False
        cell.Validation.ShowError = True
    except Exception as exc:
        print(f"  [!] Validation {address}: {exc}", flush=True)


def _unlock_addr(api: Any, address: str) -> None:
    try:
        cell = api.Range(address)
        if bool(cell.MergeCells):
            cell.MergeArea.Locked = False
        else:
            cell.Locked = False
    except Exception:
        try:
            api.Range(address).Locked = False
        except Exception:
            pass


def enforce_workbook_ui_lock(wb: Any) -> None:
    """Apply the same lockout as VBA ProtectWorkbookUi."""
    print("Enforcing workbook UI lockout...", flush=True)

    try:
        dash = wb.sheets["Dashboard"]
    except Exception as exc:
        raise RuntimeError("Dashboard sheet not found for UI lock") from exc

    d = dash.api
    _unprotect_sheet_api(d)
    try:
        d.Cells.Locked = True
        for addr in ("A3", "D3", "G3", "A5"):
            _unlock_addr(d, addr)
        for addr in ("C3", "F3", "H3", "I3", "H5"):
            _unlock_addr(d, addr)
        d.Columns(21).Locked = False  # U — hidden UNC targets
        d.Columns(21).Hidden = True
        d.Columns(13).ColumnWidth = 10  # M Size MB
        _set_list_validation(d, "C3", "AND,OR,NOT")
        _set_list_validation(d, "F3", "Over,Under")
        try:
            raw = str(d.Range("F3").Value or "").strip().upper()
            if raw in (">", "GT", "GREATER THAN"):
                d.Range("F3").Value = "Over"
            elif raw in ("<", "LT", "LESS THAN"):
                d.Range("F3").Value = "Under"
        except Exception:
            pass
        _set_list_validation(d, "H3", "Yes,No")
        _set_list_validation(d, "I3", "Yes,No,Only")
    except Exception as exc:
        print(f"  [!] Dashboard lock flags: {exc}", flush=True)
    _protect_sheet_api(d, allow_formatting_columns=True, allow_filtering=True)

    try:
        ing = wb.sheets["Ingestion"]
        iapi = ing.api
        _unprotect_sheet_api(iapi)
        try:
            iapi.Cells.Locked = True
        except Exception:
            pass
        _protect_sheet_api(iapi)
    except Exception as exc:
        print(f"  [!] Ingestion lock skipped: {exc}", flush=True)

    try:
        db = wb.sheets["Database"]
        dbapi = db.api
        _unprotect_sheet_api(dbapi)
        try:
            dbapi.Cells.Locked = True
        except Exception:
            pass
        _protect_sheet_api(dbapi)
        try:
            dbapi.Visible = XL_SHEET_VERY_HIDDEN
        except Exception as exc:
            print(f"  [!] Database VeryHidden failed: {exc}", flush=True)
    except Exception as exc:
        print(f"  [!] Database lock skipped: {exc}", flush=True)

    print("  UI lock applied (A5:G5 File Search + F3/G3 size + H5 drive).", flush=True)


def vba_project_is_locked(wb: Any) -> bool:
    """True when VBA project is locked for viewing (admin-set password)."""
    try:
        return int(wb.api.VBProject.Protection) == 1  # vbext_pp_locked
    except Exception:
        return True


def restore_excel_interactive(app: Any) -> None:
    """
    Return Excel to a normal interactive state before close/quit/reveal.

    Quitting (or Shell-reopening) while frozen/hidden often makes Windows think
    Excel crashed → "Safe mode?" on the next launch.
    """
    if app is None:
        return
    try:
        app.api.EnableEvents = True
    except Exception:
        pass
    try:
        app.enable_events = True
    except Exception:
        pass
    try:
        app.display_alerts = True
    except Exception:
        pass
    try:
        app.screen_updating = True
    except Exception:
        pass
    try:
        # Boolean False hands the bar back to Excel (do not assign a string)
        app.api.StatusBar = False
    except Exception:
        pass
    try:
        app.visible = True
    except Exception:
        pass


def soft_quit_excel_app(app: Any) -> None:
    """
    Clean Quit only — never kill the process. Call restore_excel_interactive first.
    Prefer leaving Excel open for the user when possible.
    """
    if app is None:
        return
    restore_excel_interactive(app)
    try:
        app.display_alerts = False
    except Exception:
        pass
    try:
        for book in list(app.books):
            try:
                book.save()
            except Exception:
                pass
            try:
                book.close()
            except Exception:
                pass
    except Exception:
        pass
    try:
        # COM Quit after books closed — avoids crash-recovery / safe-mode prompt
        app.api.Quit()
    except Exception:
        try:
            app.quit()
        except Exception:
            pass
    # Give Excel time to write a clean shutdown to the registry
    time.sleep(1.25)


def reopen_workbook_in_app(app: Any, path: Any) -> Any:
    """
    Close the target workbook if open on this app, restore UI, reopen so Workbook_Open fires.
    Leaves Excel running — does not Quit (avoids Safe Mode prompts).
    """
    from pathlib import Path

    target = Path(path).resolve()
    target_key = str(target).lower()
    target_name = target.name.lower()

    restore_excel_interactive(app)
    try:
        app.display_alerts = False
    except Exception:
        pass

    for book in list(app.books):
        try:
            full = str(Path(book.fullname).resolve()).lower()
        except Exception:
            full = ""
        try:
            name = book.name.lower()
        except Exception:
            name = ""
        if full == target_key or name == target_name:
            try:
                book.save()
            except Exception:
                pass
            try:
                book.close()
            except Exception:
                pass

    try:
        app.enable_events = True
        app.api.EnableEvents = True
    except Exception:
        pass

    wb = app.books.open(str(target))
    try:
        wb.activate()
    except Exception:
        pass
    restore_excel_interactive(app)
    print(f"Workbook ready: {target.name}", flush=True)
    return wb
