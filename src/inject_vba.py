"""
Close AccDB-Blank_LAN_Crawler Tool.xlsm if open, silently inject VBA from src/vba, reopen.

Pattern aligned with MasterDatabase/src/close_excel.py (xlwings silence protocol).
"""
from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

import xlwings as xw

# --- CONSTANTS ---
vbext_ct_StdModule = 1
vbext_ct_ClassModule = 2
vbext_ct_MSForm = 3
vbext_ct_Document = 100

PROJECT_ROOT = Path(__file__).resolve().parent.parent
SRC_DIR = Path(__file__).resolve().parent
DEFAULT_WORKBOOK = PROJECT_ROOT / "AccDB-Blank_LAN_Crawler Tool.xlsm"
VBA_SOURCE_DIR = SRC_DIR / "vba"

if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from workbook_security import (  # noqa: E402
    enforce_workbook_ui_lock,
    reopen_workbook_in_app,
    vba_project_is_locked,
)

# Import foundation modules before dependents.
IMPORT_PRIORITY = (
    "modPathUtil.bas",
    "modIndexCache.bas",
    "clsExcelAppState.cls",
    "clsSheetMap.cls",
    "clsFileIndexer.cls",
    "clsFileSearch.cls",
    "modEntryPoints.bas",
    "ThisWorkbook.cls",
)

# Old extracted modules — remove if still present after unlock/import.
OBSOLETE_COMPONENTS = (
    "Module1",
    "Module2",
    "Module3",
    "Module4",
)


def get_comp_type(comp) -> int:
    try:
        t = comp.Type
        if callable(t):
            t = t()
        return int(t)  # type: ignore[arg-type]
    except Exception:
        return -1


def get_code_content(file_path: str) -> str:
    with open(file_path, "r", encoding="utf-8", errors="replace") as f:
        lines = f.readlines()

    clean_lines: list[str] = []
    is_header = True
    for line in lines:
        if is_header:
            s = line.strip()
            if (
                s.startswith("VERSION")
                or s.startswith("BEGIN")
                or s.upper() == "END"
                or s.startswith("Attribute")
                or s.startswith("MultiUse")
            ):
                continue
            if not s:
                continue
            is_header = False
            clean_lines.append(line)
        else:
            clean_lines.append(line)
    return "".join(clean_lines)


def prepare_vba_import_path(file_path: str) -> tuple[str, str | None]:
    """Normalize to CRLF for VBComponents.Import. Returns (path, temp_to_delete)."""
    with open(file_path, "rb") as f:
        data = f.read()

    if b"\r\n" in data or b"\n" not in data:
        return file_path, None

    text = data.decode("utf-8", errors="replace")
    normalized = text.replace("\r\n", "\n").replace("\n", "\r\n")
    fd, temp_path = tempfile.mkstemp(
        suffix=os.path.splitext(file_path)[1],
        prefix="vba_import_",
    )
    os.close(fd)
    with open(temp_path, "wb") as f:
        f.write(normalized.encode("utf-8"))
    return temp_path, temp_path


def build_import_list(source_dir: Path) -> list[Path]:
    files: list[Path] = []
    seen: set[str] = set()

    for name in IMPORT_PRIORITY:
        path = source_dir / name
        if path.is_file():
            files.append(path)
            seen.add(name.lower())

    for path in sorted(source_dir.iterdir()):
        if path.suffix.lower() not in {".bas", ".cls", ".frm"}:
            continue
        if path.name.lower() in seen:
            continue
        files.append(path)
    return files


def _silence_app(app) -> None:
    """Freeze Excel UI so open/close/save does not flicker."""
    try:
        app.screen_updating = False
    except Exception:
        pass
    try:
        app.display_alerts = False
    except Exception:
        pass
    try:
        app.enable_events = False
    except Exception:
        pass
    try:
        app.api.EnableEvents = False
    except Exception:
        pass


def _unsilence_app(app) -> None:
    try:
        app.enable_events = True
    except Exception:
        pass
    try:
        app.api.EnableEvents = True
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


def close_if_open(target_file: Path) -> None:
    target_key = str(target_file.resolve()).lower()
    target_name = target_file.name.lower()

    for existing_app in list(xw.apps):
        for book in list(existing_app.books):
            try:
                full = str(Path(book.fullname).resolve()).lower()
            except Exception:
                full = ""
            try:
                name = book.name.lower()
            except Exception:
                name = ""

            if full == target_key or name == target_name:
                print(f"Closing existing instance of {book.name}...")
                _silence_app(existing_app)
                try:
                    book.save()
                except Exception:
                    pass
                try:
                    book.close()
                except Exception:
                    pass
                # Only quit this Excel if it has no books left — restore UI first
                try:
                    if len(existing_app.books) == 0:
                        from workbook_security import soft_quit_excel_app

                        soft_quit_excel_app(existing_app)
                except Exception:
                    pass


def remove_component(vb_project, name: str) -> None:
    try:
        comp = vb_project.VBComponents(name)
    except Exception:
        return

    c_type = get_comp_type(comp)
    if c_type == vbext_ct_Document:
        return

    safe_name = name[:20] + "_DEL"
    try:
        ghost = vb_project.VBComponents(safe_name)
        vb_project.VBComponents.Remove(ghost)
    except Exception:
        pass

    try:
        comp.Name = safe_name
        vb_project.VBComponents.Remove(comp)
        print(f"  > Removed obsolete: {name}")
    except Exception as exc:
        print(f"  [!] Could not remove {name}: {exc}")


def import_component(vb_project, file_path: Path) -> None:
    name = file_path.stem
    ext = file_path.suffix.lower()

    try:
        comp = vb_project.VBComponents(name)
    except Exception:
        comp = None

    c_type = get_comp_type(comp) if comp else -1

    if comp and c_type == vbext_ct_Document and ext == ".cls":
        print(f"  > Injecting code into document object: {name}")
        code_text = get_code_content(str(file_path))
        if comp.CodeModule.CountOfLines > 0:
            comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
        if code_text.strip():
            comp.CodeModule.AddFromString(code_text)
        return

    if comp:
        print(f"  > Updating component: {name}")
        safe_name = name[:20] + "_DEL"
        try:
            ghost = vb_project.VBComponents(safe_name)
            vb_project.VBComponents.Remove(ghost)
        except Exception:
            pass
        comp.Name = safe_name
        vb_project.VBComponents.Remove(comp)
    else:
        print(f"  > Importing new component: {name}")

    import_path, temp_path = prepare_vba_import_path(str(file_path))
    try:
        vb_project.VBComponents.Import(import_path)
    finally:
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)


def inject_vba(target_file: Path, source_dir: Path, reopen: bool = True) -> None:
    if not target_file.is_file():
        raise FileNotFoundError(f"Workbook not found: {target_file}")
    if not source_dir.is_dir():
        raise FileNotFoundError(f"VBA source folder not found: {source_dir}")

    print(f"Target Excel: {target_file}")
    print(f"VBA source:   {source_dir}")

    close_if_open(target_file)

    app = None
    try:
        print("Opening Excel (hidden, UI frozen)...")
        app = xw.App(visible=False, add_book=False)
        _silence_app(app)
        try:
            app.visible = False
        except Exception:
            pass

        wb = app.books.open(str(target_file))

        if vba_project_is_locked(wb):
            raise RuntimeError(
                "VBA project is password-protected. Unlock it once in the VBE "
                "(Tools > VBAProject Properties > Protection), save the workbook, "
                "then re-run inject. Sheet data import does not need this unlock."
            )

        vb_project = wb.api.VBProject

        print("Removing obsolete modules (if present)...")
        for obsolete in OBSOLETE_COMPONENTS:
            remove_component(vb_project, obsolete)

        print(f"Importing VBA modules from '{source_dir}'...")
        for path in build_import_list(source_dir):
            import_component(vb_project, path)

        print("Applying UI lockout...")
        enforce_workbook_ui_lock(wb)

        print("Saving changes...")
        wb.save()

        if reopen:
            saved_path = str(target_file.resolve())
            print("Reopening workbook in same Excel (Workbook_Open / cache warm)...")
            reopen_workbook_in_app(app, saved_path)
            app = None  # leave Excel running for the user — do not Quit
            print("Update complete. Excel is open.")
        else:
            from workbook_security import soft_quit_excel_app  # noqa: E402

            try:
                wb.close()
            except Exception:
                pass
            soft_quit_excel_app(app)
            app = None
            print("Update complete. Excel closed cleanly.")
    finally:
        if app is not None:
            try:
                from workbook_security import soft_quit_excel_app  # noqa: E402

                soft_quit_excel_app(app)
            except Exception:
                try:
                    app.quit()
                except Exception:
                    pass

def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    reopen = True
    if "--no-reopen" in args:
        reopen = False
        args.remove("--no-reopen")

    if args:
        target = Path(args[0])
    else:
        target = DEFAULT_WORKBOOK

    if not target.is_absolute():
        target = (PROJECT_ROOT / target).resolve()

    try:
        inject_vba(target, VBA_SOURCE_DIR, reopen=reopen)
        return 0
    except Exception as exc:
        print(f"CRITICAL FAILURE: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
