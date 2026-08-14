"""
Unlock AccDB-Blank (or another .xlsm) for Dashboard layout editing.

Unprotects sheets, unlocks ALL cells, sets layout-edit mode so WarmIndexCache
will not re-lock, leaves Excel open.
Does not touch the VBA project password.
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path
from typing import Any

_SRC = Path(__file__).resolve().parent
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from workbook_security import _unprotect_sheet_api  # noqa: E402

XL_SHEET_VISIBLE = -1
DEFAULT_WB = Path(__file__).resolve().parent.parent / "AccDB-Blank_LAN_Crawler Tool.xlsm"


def _com_unlock_all(wb: Any) -> None:
    for name in ("Dashboard", "Ingestion", "Database"):
        try:
            ws = wb.sheets[name]
        except Exception:
            print(f"  [!] Sheet missing: {name}", flush=True)
            continue
        api = ws.api
        _unprotect_sheet_api(api)
        try:
            api.Visible = XL_SHEET_VISIBLE
        except Exception:
            pass
        try:
            # Unlock every cell — labels outside B3:F3 were still Locked under UI protect
            api.Cells.Locked = False
        except Exception as exc:
            print(f"  [!] Unlock cells on {name}: {exc}", flush=True)
        try:
            api.EnableSelection = 0  # xlNoRestrictions
        except Exception:
            pass
        # Ensure sheet is not protected
        _unprotect_sheet_api(api)
        print(f"  Unprotected + all cells unlocked: {name}", flush=True)


def unlock_workbook(workbook: Path) -> None:
    import xlwings as xw

    workbook = Path(workbook)
    if not workbook.is_file():
        raise FileNotFoundError(workbook)

    saved = str(workbook.resolve())
    print(f"Unlocking for edit: {saved}", flush=True)

    app = xw.apps.active
    if app is None:
        app = xw.App(visible=True, add_book=False)

    app.visible = True
    app.display_alerts = False
    try:
        app.enable_events = False
    except Exception:
        try:
            app.api.EnableEvents = False
        except Exception:
            pass

    wb = None
    target = workbook.resolve()
    for b in app.books:
        try:
            if Path(b.fullname).resolve() == target:
                wb = b
                break
        except Exception:
            pass
    if wb is None:
        wb = app.books.open(saved)

    # Prefer VBA UnlockForLayoutEdit (sets gLayoutEditMode so warm won't re-lock)
    ran_vba = False
    try:
        app.api.Run("UnlockForLayoutEdit")
        ran_vba = True
        print("  Ran UnlockForLayoutEdit (layout edit mode ON).", flush=True)
    except Exception as exc:
        print(f"  UnlockForLayoutEdit not available yet ({exc}).", flush=True)
        print("  Applying COM unlock; inject VBA then re-run this task to stop auto re-lock.", flush=True)

    _com_unlock_all(wb)

    # WarmIndexCache often fires ~2s after open and re-locks — clear again after that window
    if not ran_vba:
        print("  Waiting 3s in case WarmIndexCache re-locks, then unlocking again...", flush=True)
        time.sleep(3.0)
        _com_unlock_all(wb)

    try:
        wb.sheets["Dashboard"].activate()
    except Exception:
        pass

    try:
        app.enable_events = True
    except Exception:
        try:
            app.api.EnableEvents = True
        except Exception:
            pass
    app.display_alerts = True

    print(
        "Done. All Dashboard cells should be editable now.\n"
        "Edit labels / merge as needed. When finished: Import VBA, then RelockAfterLayoutEdit "
        "(or just re-open / search — inject reapplies UI lock).",
        flush=True,
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Unlock LAN Search workbook UI for layout edits")
    parser.add_argument(
        "workbook",
        nargs="?",
        default=str(DEFAULT_WB),
        help="Path to .xlsm (default: AccDB-Blank_LAN_Crawler Tool.xlsm)",
    )
    args = parser.parse_args(argv)
    try:
        unlock_workbook(Path(args.workbook))
    except Exception as exc:  # noqa: BLE001
        print(f"Unlock failed: {exc}", flush=True)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
