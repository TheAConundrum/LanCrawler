"""Pack dist/Lan_Search_Tool.zip with offline wheels (packer's pip / JFrog)."""
from __future__ import annotations

import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
SRC_DIR = PROJECT_ROOT / "src"
REQ = PROJECT_ROOT / "requirements-crawler.txt"
WORKBOOK = PROJECT_ROOT / "Lan_Search_Tool.xlsm"
WHEELS_DIR = PROJECT_ROOT / "vendor" / "wheels"
DIST_DIR = PROJECT_ROOT / "dist"
ZIP_NAME = "Lan_Search_Tool.zip"
ZIP_ROOT = "Lan_Search_Tool"


PACK_PYTHON_VERSIONS = ("3.10", "3.11", "3.12", "3.13", "3.14")
PACK_PLATFORMS = ("win_amd64", "win32")


def _cp_tag_to_version(tag: str) -> str | None:
    if not tag.startswith("cp") or not tag[2:].isdigit():
        return None
    digits = tag[2:]
    if len(digits) == 2:
        return f"{digits[0]}.{digits[1]}"
    if len(digits) == 3:
        return f"{digits[0]}.{digits[1:]}"
    return None


def _packed_python_versions() -> list[str]:
    found: set[str] = set()
    if not WHEELS_DIR.is_dir():
        return []
    for path in WHEELS_DIR.glob("*.whl"):
        for part in path.stem.split("-"):
            ver = _cp_tag_to_version(part)
            if ver:
                found.add(ver)
    return sorted(found, key=lambda s: [int(x) for x in s.split(".")])


def _readme_txt() -> str:
    packed = _packed_python_versions()
    ver_list = ", ".join(packed) if packed else f"{sys.version_info.major}.{sys.version_info.minor}"
    text = f"""LAN Search Tool
================

This kit indexes a folder or drive and saves the results in an Access
database next to Lan_Search_Tool.xlsm. Open that workbook in Excel to
search the files.

This copy includes offline Python packages for:
  Python {ver_list}  (Windows 64-bit and 32-bit)

Check your version with:  py --version   or   python --version


HOW TO RUN
----------

1. Unzip the whole folder. Do not move run_crawl.cmd out by itself.
   Keep it next to Lan_Search_Tool.xlsm, src, and vendor.

2. Double-click run_crawl.cmd
   The first run may take a minute while it sets up temporary Python
   packages from the vendor folder (it does not download from the
   internet).

3. When the window opens, click Browse and pick the folder or drive
   you want to index, then click Start crawl.

4. If you already crawled before, you will be asked:

      Do you want to add to the DB taken on <date>,
      or delete it and start fresh?

   - Add to existing  - keep the previous index and add this folder
   - Start fresh      - erase the old index and use only this crawl
   - Cancel           - do nothing

5. Wait until the crawl finishes. You can add another folder later by
   running run_crawl.cmd again and choosing Add to existing.

6. At the end you will be asked:

      Do you want to delete the temporary python packages? [Y/N]

   Y removes the temporary packages so they do not stay on this PC.
   N keeps them so the next crawl starts faster.

7. Open Lan_Search_Tool.xlsm in Excel to search. Close Excel before
   you crawl again if the database is in use.


WHAT YOU NEED ON THIS PC
------------------------

- Python {ver_list} with "py" or "python" on PATH
- Microsoft Excel
- Microsoft Access Database Engine (ACE), same 32/64-bit as **Python**
  (if 32-bit Office is installed, install 64-bit ACE with /quiet)

You do not need to install Python packages yourself.


WHERE THE INDEX IS SAVED
------------------------

Workbook:  Lan_Search_Tool.xlsm  (in this folder)
Database:  DB\\SearchIndex-M-D-YYYY.accdb  (created beside the workbook)

The date in the file name is when that database was first created.
Later crawls add more folders into the same file unless you choose
Start fresh.


IF SOMETHING GOES WRONG
-----------------------

- "python is not recognized"
  Install Python ({ver_list}) and check "Add python.exe to PATH", or
  use the Windows py launcher.

- "Failed to install packages from vendor\\wheels"
  Your Python version or 32/64-bit does not match a wheel in vendor\\wheels.
  The command window lists this PC's Python and the packed wheel names
  (look for cp310, cp311, cp312, cp313, cp314 and win_amd64 vs win32).

- AccDB / "Data source name not found" / IM002
  The crawl worked, but Windows has no Access ODBC driver that matches
  this Python (often 64-bit Python + 32-bit Office).

  Install Microsoft Access Database Engine 2016 Redistributable,
  same 32/64-bit as Python (not as Office). For 64-bit Python with
  32-bit Office already installed, from an elevated command prompt:

    AccessDatabaseEngine_X64.exe /quiet

  Then close Excel and run run_crawl.cmd again (choose Add to existing
  if this folder was already crawled, or Start fresh).

  Also close Lan_Search_Tool.xlsm before crawling if the database is open.

- Keep this folder together. Do not run run_crawl.cmd from a copy
  that is missing src, vendor, or Lan_Search_Tool.xlsm.
"""
    return text.replace("\n", "\r\n")


def _pip_download() -> None:
    if WHEELS_DIR.exists():
        shutil.rmtree(WHEELS_DIR)
    WHEELS_DIR.mkdir(parents=True, exist_ok=True)
    failed: list[str] = []
    print("Downloading wheels with packer pip (JFrog / configured index)...")
    for ver in PACK_PYTHON_VERSIONS:
        abi = f"cp{ver.replace('.', '')}"
        for platform in PACK_PLATFORMS:
            label = f"{ver} {platform}"
            cmd = [
                sys.executable,
                "-m",
                "pip",
                "download",
                "--only-binary",
                ":all:",
                "--python-version",
                ver,
                "--platform",
                platform,
                "--implementation",
                "cp",
                "--abi",
                abi,
                "-r",
                str(REQ),
                "-d",
                str(WHEELS_DIR),
            ]
            print(f"  {label}: {' '.join(cmd)}")
            result = subprocess.run(cmd)
            if result.returncode != 0:
                failed.append(label)
                print(f"  Warning: no wheels for {label}")
    if not any(WHEELS_DIR.glob("*.whl")):
        raise SystemExit("vendor/wheels is empty after pip download")
    if failed:
        print("Some Python/platform combos had no wheels: " + ", ".join(failed))
    print("Packed Python versions: " + ", ".join(_packed_python_versions()))


def _add_tree(zf: zipfile.ZipFile, src: Path, arc_prefix: str) -> None:
    for path in src.rglob("*"):
        if path.is_dir():
            continue
        if path.suffix.lower() != ".py":
            continue
        if path.name == "__pycache__" or "__pycache__" in path.parts:
            continue
        rel = path.relative_to(src)
        zf.write(path, f"{arc_prefix}/{rel.as_posix()}")


def pack() -> Path:
    if not WORKBOOK.is_file():
        raise SystemExit(f"Workbook not found: {WORKBOOK}")
    if not REQ.is_file():
        raise SystemExit(f"Missing {REQ}")
    _pip_download()
    if not any(WHEELS_DIR.iterdir()):
        raise SystemExit("vendor/wheels is empty after pip download")

    DIST_DIR.mkdir(parents=True, exist_ok=True)
    zip_path = DIST_DIR / ZIP_NAME
    if zip_path.exists():
        zip_path.unlink()

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.write(PROJECT_ROOT / "run_crawl.cmd", f"{ZIP_ROOT}/run_crawl.cmd")
        zf.write(REQ, f"{ZIP_ROOT}/requirements-crawler.txt")
        zf.write(WORKBOOK, f"{ZIP_ROOT}/Lan_Search_Tool.xlsm")
        zf.writestr(f"{ZIP_ROOT}/ReadMe.txt", _readme_txt())
        _add_tree(zf, SRC_DIR / "crawler", f"{ZIP_ROOT}/src/crawler")
        zf.write(SRC_DIR / "workbook_security.py", f"{ZIP_ROOT}/src/workbook_security.py")
        zf.writestr(f"{ZIP_ROOT}/src/__init__.py", "")
        for wheel in sorted(WHEELS_DIR.iterdir()):
            if wheel.is_file():
                zf.write(wheel, f"{ZIP_ROOT}/vendor/wheels/{wheel.name}")

    print(f"Wrote {zip_path}")
    packed = ", ".join(_packed_python_versions()) or "(none)"
    print(f"Python wheels tagged for: {packed}")
    return zip_path


def main() -> int:
    try:
        pack()
    except subprocess.CalledProcessError as exc:
        print(f"pip download failed (need JFrog / configured pip index): {exc}", flush=True)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
