"""
Workbook UI lock helpers for LAN Search Tool.

Mirrors clsSheetMap.ProtectWorkbookUi:
  - Dashboard free-text: Key Word B/D and Size MB F (rows 3–5)
  - Dashboard dropdowns (unlocked + list validation): op C, size op E,
    Flexible H4, Search folders I4, Drive H5
  - Ingestion: fully locked
  - Database: locked + xlSheetVeryHidden
  - EnableSelection = xlNoRestrictions (no protect-popup on select)

VBA project password is NOT set here — lock the VBE project once as admin
(Tools → VBAProject Properties → Protection). Sheet data import does not need
VBA unlocked; inject_vba.py does and will fail clearly if the project is locked.
"""

from __future__ import annotations

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


def _protect_sheet_api(api: Any, *, allow_formatting_columns: bool = False) -> None:
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
            AllowSorting=False,
            AllowFiltering=False,
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
        for col in (2, 4, 6):  # B, D, F — keywords + Size MB
            d.Range(d.Cells(3, col), d.Cells(5, col)).Locked = False
        for col in (3, 5):  # C op, E size op — dropdown select
            d.Range(d.Cells(3, col), d.Cells(5, col)).Locked = False
        for addr in ("H4", "I4", "H5"):
            d.Range(addr).Locked = False
        d.Columns(21).Locked = False  # U — hidden UNC targets
        d.Columns(21).Hidden = True
        for row in range(3, 6):
            _set_list_validation(d, f"C{row}", "AND,OR,NOT")
            _set_list_validation(d, f"E{row}", "<,>")
        _set_list_validation(d, "H4", "Yes,No")
        _set_list_validation(d, "I4", "Yes,No,Only")
    except Exception as exc:
        print(f"  [!] Dashboard lock flags: {exc}", flush=True)
    # Dashboard: allow column resize (G/K/O+); VBA snaps A:F, H:J, L:N back
    _protect_sheet_api(d, allow_formatting_columns=True)

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

    print("  UI lock applied (keywords+MB+dropdowns; Dashboard cols resizable except chrome).", flush=True)


def vba_project_is_locked(wb: Any) -> bool:
    """True when VBA project is locked for viewing (admin-set password)."""
    try:
        return int(wb.api.VBProject.Protection) == 1  # vbext_pp_locked
    except Exception:
        # Cannot read project — treat as locked / inaccessible
        return True
