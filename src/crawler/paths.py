"""Path helpers: UNC resolve, size MB, parent walks."""
from __future__ import annotations

import ctypes
import string
from ctypes import wintypes
from pathlib import Path, PureWindowsPath

BYTES_PER_MB = 1048576.0
MIN_SIZE_MB = 0.01

ERROR_SUCCESS = 0
ERROR_MORE_DATA = 234


def bytes_to_size_mb(size_bytes: float | int) -> float:
    if size_bytes is None or size_bytes <= 0:
        return MIN_SIZE_MB
    mb = float(size_bytes) / BYTES_PER_MB
    if mb < MIN_SIZE_MB:
        return MIN_SIZE_MB
    return round(mb, 2)


def normalize_slashes(path: str) -> str:
    return path.replace("/", "\\").strip()


def strip_trailing_slash(path: str) -> str:
    p = normalize_slashes(path)
    while len(p) > 3 and p.endswith("\\"):
        p = p[:-1]
    # UNC root like \\server\share should keep no trailing slash for keys
    if p.endswith("\\") and not p.startswith("\\\\"):
        p = p[:-1]
    if p.startswith("\\\\") and p.endswith("\\"):
        p = p[:-1]
    return p


def parent_folder_path(full_path: str) -> str:
    """Parent path without trailing slash (empty if none)."""
    p = strip_trailing_slash(full_path)
    if not p:
        return ""
    parent = str(PureWindowsPath(p).parent)
    if parent in {".", ""}:
        return ""
    # PureWindowsPath("\\server\share").parent can be "\\"
    if parent == "\\":
        return ""
    return strip_trailing_slash(parent)


def get_remote_name_for_drive(drive_with_colon: str) -> str:
    """Map 'O:' -> '\\\\server\\share' via WNetGetConnectionA."""
    drive = drive_with_colon.upper().rstrip("\\")
    if len(drive) == 1:
        drive = drive + ":"
    if len(drive) != 2 or drive[1] != ":":
        return ""

    mpr = ctypes.WinDLL("mpr")
    buf_len = wintypes.DWORD(0)
    local = ctypes.c_char_p(drive.encode("ascii", errors="ignore"))

    # First call to get required size
    remote = ctypes.create_string_buffer(1)
    err = mpr.WNetGetConnectionA(local, remote, ctypes.byref(buf_len))
    if err not in (ERROR_SUCCESS, ERROR_MORE_DATA) or buf_len.value <= 0:
        return ""

    remote = ctypes.create_string_buffer(buf_len.value)
    err = mpr.WNetGetConnectionA(local, remote, ctypes.byref(buf_len))
    if err != ERROR_SUCCESS:
        return ""
    return remote.value.decode("ascii", errors="ignore").rstrip("\x00")


def build_drive_map(paths: list[str] | None = None) -> dict[str, str]:
    """Cache drive-letter -> UNC root for common mapped drives."""
    mapping: dict[str, str] = {}
    drives: set[str] = set()
    if paths:
        for p in paths:
            p = normalize_slashes(p)
            if len(p) >= 2 and p[1] == ":":
                drives.add(p[:2].upper())
    else:
        for letter in string.ascii_uppercase:
            drives.add(f"{letter}:")

    for drive in sorted(drives):
        remote = get_remote_name_for_drive(drive)
        if remote:
            mapping[drive] = strip_trailing_slash(remote)
    return mapping


def to_unc_path(path: str, drive_map: dict[str, str] | None = None, unc_root_override: str = "") -> str:
    trimmed = normalize_slashes(path)
    if not trimmed:
        return ""

    if trimmed.startswith("\\\\"):
        return strip_trailing_slash(trimmed)

    if len(trimmed) >= 3 and trimmed[1] == ":" and trimmed[2] in "\\/":
        drive = trimmed[:2].upper()
        rest = trimmed[3:].lstrip("\\")
        remote = ""
        if drive_map and drive in drive_map:
            remote = drive_map[drive]
        else:
            remote = get_remote_name_for_drive(drive)
        if remote:
            return strip_trailing_slash(str(PureWindowsPath(remote) / rest)) if rest else strip_trailing_slash(remote)
        if unc_root_override:
            root = strip_trailing_slash(unc_root_override)
            return strip_trailing_slash(str(PureWindowsPath(root) / rest)) if rest else root

    return strip_trailing_slash(trimmed)


def path_starts_with_root(file_path: str, root_path: str) -> bool:
    f = normalize_slashes(file_path).lower()
    r = strip_trailing_slash(root_path).lower() + "\\"
    if not r.strip("\\"):
        return False
    bare = r[:-1]
    return f.startswith(r) or f == bare


def ensure_crawl_start(path: str) -> Path:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Folder not found: {path}")
    if not p.is_dir():
        raise NotADirectoryError(f"Not a folder: {path}")
    return p
