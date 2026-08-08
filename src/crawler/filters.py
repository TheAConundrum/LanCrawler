"""File/folder filters mirrored from VBA modPathUtil."""
from __future__ import annotations

import re
from pathlib import PureWindowsPath

ALLOWED_INDEX_EXTS = frozenset(
    {
        "PDF",
        "DOC",
        "DOCX",
        "XLS",
        "XLSX",
        "XLSM",
        "XLSB",
        "PPT",
        "PPTX",
        "TXT",
        "RTF",
        "CSV",
        "ODT",
        "ODS",
        "ODP",
        "MSG",
        "EML",
        "JPG",
        "JPEG",
        "PNG",
        "GIF",
        "TIF",
        "TIFF",
        "BMP",
        "WEBP",
        "ZIP",
        "7Z",
        "RAR",
        "TAR",
        "GZ",
        "TGZ",
        "BZ2",
        "XZ",
        "CAB",
        "MP4",
        "MOV",
        "AVI",
        "MKV",
        "WMV",
        "WEBM",
        "M4V",
        "MPG",
        "MPEG",
        "FLV",
        "3GP",
        "TS",
        "MTS",
        "M2TS",
    }
)

JUNK_FOLDERS = frozenset(
    {
        "$RECYCLE.BIN",
        "SYSTEM VOLUME INFORMATION",
        "THUMBS.DB",
        ".GIT",
        ".SVN",
        ".HG",
        "__PYCACHE__",
        "NODE_MODULES",
        ".VS",
        ".IDEA",
        "RECYCLER",
    }
)

JUNK_FILE_NAMES = frozenset(
    {
        "THUMBS.DB",
        "DESKTOP.INI",
        ".DS_STORE",
    }
)

JUNK_FILE_EXTS = frozenset({"TMP", "TEMP", "BAK", "LOG", "LCK", "PART"})

_DIGITS = re.compile(r"^\d+$")


def _base_name_upper(file_name: str) -> str:
    return PureWindowsPath(file_name.replace("/", "\\")).name.upper()


def extension_only(file_name: str) -> str:
    base = _base_name_upper(file_name)
    if not base:
        return ""
    dot = base.rfind(".")
    if dot <= 0 or dot == len(base) - 1:
        return ""
    return base[dot + 1 :]


def is_digits_only(text: str) -> bool:
    return bool(text) and _DIGITS.match(text) is not None


def is_split_archive_volume(file_name: str, ext: str | None = None) -> bool:
    """True for multi-volume pieces when the parent archive is enough."""
    base = _base_name_upper(file_name)
    ext_u = (ext or extension_only(file_name)).upper()

    # Classic RAR volumes: name.r00 .. name.r100
    if ext_u.startswith("R") and 3 <= len(ext_u) <= 4 and is_digits_only(ext_u[1:]):
        return True

    # Split ZIP volumes: name.z01 .. name.z99
    if ext_u.startswith("Z") and len(ext_u) == 3 and is_digits_only(ext_u[1:]):
        return True

    # name.7z.001 / name.zip.001 / name.rar.001
    if is_digits_only(ext_u) and len(ext_u) in (2, 3):
        stem = base[: -(len(ext_u) + 1)]
        if stem.endswith(".7Z") or stem.endswith(".ZIP") or stem.endswith(".RAR"):
            return True

    # name.part2.rar — keep part1 / part01 only
    if ext_u in {"RAR", "ZIP", "7Z"}:
        p = base.rfind(".PART")
        if p >= 0:
            part_tok = base[p + 5 :]
            dot = part_tok.rfind(".")
            if dot > 0:
                part_num = part_tok[:dot]
                if is_digits_only(part_num) and int(part_num) != 1:
                    return True

    return False


def is_junk_folder_name(folder_name: str) -> bool:
    return folder_name.strip().upper() in JUNK_FOLDERS


def is_junk_file_name(file_name: str) -> bool:
    n = file_name.strip()
    if not n:
        return True
    if n.startswith("~$"):
        return True
    ext = extension_only(n)
    upper = n.upper()
    if upper in JUNK_FILE_NAMES:
        return True
    if ext in JUNK_FILE_EXTS:
        return True
    return False


def is_allowed_index_extension(file_name: str) -> bool:
    ext = extension_only(file_name)
    if not ext or ext not in ALLOWED_INDEX_EXTS:
        return False
    if is_split_archive_volume(file_name, ext):
        return False
    return True
