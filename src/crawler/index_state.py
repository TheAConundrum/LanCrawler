"""Previous-index snapshots for incremental LAN crawls.

Skip a full scandir when a folder's directory mtime is unchanged, but still:
- re-stat known indexed files (in-place resave / size change)
- recurse into known child folders (new/changed files in a subfolder)
"""
from __future__ import annotations

from dataclasses import dataclass, field

try:
    from .paths import BYTES_PER_MB, parent_folder_path, strip_trailing_slash
except ImportError:
    from crawler.paths import BYTES_PER_MB, parent_folder_path, strip_trailing_slash


def mtime_key(ts: float) -> int:
    """1-second resolution so SMB/NTFS jitter does not force a rescan."""
    if ts <= 0:
        return 0
    return int(ts)


def folder_mtime_unchanged(disk_mtime: float, stored_mtime: float) -> bool:
    stored = mtime_key(stored_mtime)
    if stored <= 0:
        return False
    return mtime_key(disk_mtime) == stored


def row_size_bytes(size_bytes: int, size_mb: float) -> int:
    if size_bytes > 0:
        return int(size_bytes)
    try:
        return int(round(float(size_mb) * BYTES_PER_MB))
    except (TypeError, ValueError):
        return 0


def norm_folder_key(path: str) -> str:
    return strip_trailing_slash(path).lower()


@dataclass
class FileSnapshot:
    path: str
    file_date: str
    size_mb: float
    mtime_ts: float = 0.0
    size_bytes: int = 0


@dataclass
class FolderSnapshot:
    unc: str
    dir_mtime: float = 0.0
    local_bytes: int = 0
    created_date: str = ""
    files: list[FileSnapshot] = field(default_factory=list)
    child_uncs: list[str] = field(default_factory=list)


@dataclass
class PreviousIndex:
    folders: dict[str, FolderSnapshot] = field(default_factory=dict)
    has_meta: bool = False

    def get(self, unc: str) -> FolderSnapshot | None:
        return self.folders.get(norm_folder_key(unc))


@dataclass
class FolderMetaRow:
    folder_path: str
    dir_mtime: float
    local_bytes: int
    created_date: str = ""


def build_previous_index(
    *,
    file_rows: list[tuple[str, str, float, str]],
    meta_rows: list[FolderMetaRow] | None = None,
) -> PreviousIndex:
    """
    file_rows: (FilePath, FileDate, SizeMB, EntryType)
    """
    files_by_parent: dict[str, list[FileSnapshot]] = {}
    folder_uncs: list[str] = []

    for path, file_date, size_mb, entry_type in file_rows:
        path = strip_trailing_slash(path)
        if not path:
            continue
        et = (entry_type or "FILE").strip().upper()
        if et == "FOLDER":
            folder_uncs.append(path)
            continue
        parent = parent_folder_path(path)
        key = norm_folder_key(parent) if parent else ""
        files_by_parent.setdefault(key, []).append(
            FileSnapshot(
                path=path,
                file_date=file_date or "",
                size_mb=float(size_mb or 0),
                size_bytes=row_size_bytes(0, float(size_mb or 0)),
            )
        )

    children_by_parent: dict[str, list[str]] = {}
    for folder_unc in folder_uncs:
        parent = parent_folder_path(folder_unc)
        pkey = norm_folder_key(parent) if parent else ""
        children_by_parent.setdefault(pkey, []).append(folder_unc)

    meta_by_key: dict[str, FolderMetaRow] = {}
    for meta in meta_rows or []:
        meta_by_key[norm_folder_key(meta.folder_path)] = meta

    prev = PreviousIndex(has_meta=bool(meta_rows))
    seen_keys: set[str] = set()
    for folder_unc in folder_uncs:
        key = norm_folder_key(folder_unc)
        if key in seen_keys:
            continue
        seen_keys.add(key)
        meta = meta_by_key.get(key)
        prev.folders[key] = FolderSnapshot(
            unc=folder_unc,
            dir_mtime=float(meta.dir_mtime) if meta else 0.0,
            local_bytes=int(meta.local_bytes) if meta else 0,
            created_date=(meta.created_date if meta else "") or "",
            files=files_by_parent.get(key, []),
            child_uncs=list(children_by_parent.get(key, [])),
        )
    return prev
