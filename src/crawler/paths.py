"""Path helpers: UNC resolve, size MB, parent walks."""
from __future__ import annotations

import ctypes
import os
import string
import threading
from ctypes import wintypes
from pathlib import Path, PureWindowsPath

DRIVE_REMOTE = 4

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


def _root_is_reachable(root: str, timeout_sec: float = 20.0) -> bool:
    """True if the drive root answers before timeout (disconnected maps can hang)."""
    box: dict[str, bool] = {"ok": False}

    def _check() -> None:
        try:
            box["ok"] = os.path.isdir(root)
        except OSError:
            box["ok"] = False

    t = threading.Thread(target=_check, daemon=True)
    t.start()
    t.join(timeout_sec)
    return bool(box["ok"])


def list_mapped_network_drives() -> list[tuple[str, str, str]]:
    """Mapped network drives: (letter like 'S:', root 'S:\\', UNC). Skips disconnected maps."""
    if os.name != "nt":
        return []

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    get_drive_type = kernel32.GetDriveTypeW
    get_drive_type.argtypes = [wintypes.LPCWSTR]
    get_drive_type.restype = wintypes.UINT

    found: list[tuple[str, str, str]] = []
    for letter in string.ascii_uppercase:
        root = f"{letter}:\\"
        try:
            dtype = int(get_drive_type(root))
        except Exception:
            continue
        if dtype != DRIVE_REMOTE:
            continue
        letter_colon = f"{letter}:"
        if not _root_is_reachable(root):
            continue
        unc = get_remote_name_for_drive(letter_colon)
        found.append((letter_colon, root, strip_trailing_slash(unc) if unc else letter_colon))
    return found


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
