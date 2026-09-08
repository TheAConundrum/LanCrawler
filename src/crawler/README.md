# LAN Search Tool — Parallel Python Crawler

Faster LAN indexing than the VBA `FileSystemObject` crawl. Walks folders with a **thread pool** + `os.scandir`, applies the same allowlist / junk / split-archive rules as `modPathUtil`, and writes an exclusive index: onboard `Database!tblFiles` (small) or Access AccDB (large).

## Why it’s faster

LAN crawls are dominated by waiting on directory listings. Multiple workers overlap those waits. VBA cannot do this cleanly inside Excel.

## Requirements

- Python 3.10+ on Windows
- Stdlib for crawl/CSV
- AccDB + Excel import: `pip install -r requirements-crawler.txt` (`xlwings`, `pyodbc`, `pywin32`)
- Microsoft Access Database Engine (ACE) matching Office bitness (for AccDB create/query)

## Quick start (easiest in Cursor)

**Run Task:** `Crawl LAN folder (picker)`  
or double-click `run_crawl.cmd` (folder picker GUI).

**All mapped drives:** double-click `SummitLANScan.cmd` (same GUI: workers + update mode, then live progress). Each launcher closes its command window as soon as the GUI opens.

A GUI asks for workers and update mode (and a folder for `run_crawl.cmd`; workbook defaults to `Lan_Search_Tool.xlsm`). A **live progress window** shows folders done, full listings vs quick skips, files, pending, and errors.

**Quick update** (default when an AccDB already exists): skip a full `scandir` on folders whose directory timestamp is unchanged. Still re-stats known indexed files (so a resave to a new date is captured) and still walks known subfolders (so a file added in a child folder is captured). Other crawl roots already in the AccDB are kept.

**Full recrawl**: re-list every folder under the selected path (ReplaceRoot merge; other drives stay).

**Start fresh**: deletes `DB\SearchIndex-*.accdb` and `LAN_Search_Index.accdb`, then creates `SearchIndex-{today}.accdb`.

The first crawl after this upgrade has no folder timestamps yet, so it is a full listing. That run writes `tblFolderMeta`; later weekly runs can be quick.

If the AccDB is already gone, no prompt — a new dated file is created. AccDB crawls write the index in memory → AccDB only (no `crawl_output` CSV).

Keep `Blank_LAN_Crawler Tool.xlsm` as the stable non-AccDB template; AccDB work uses `Lan_Search_Tool.xlsm`.

UNC also works:

```bat
python -m crawler "\\fileserver\share$\Some Folder" -w 24
```

### Common flags

| Flag | Meaning |
|------|---------|
| `-w` / `--workers` | Parallel folder workers (default **16**; try 8–32) |
| `--full` | Re-list every folder (skip quick update) |
| `-o path.csv` | Optional CSV dump (not written unless this flag is set) |
| `--accdb path.accdb` | Write AccDB `tblFiles` |
| `--target Auto\|Workbook\|AccDB` | Exclusive index mode with `--import-excel` |
| `--fresh` | Delete existing AccDB files, then write a new snapshot (CLI; no GUI prompt) |
| `--unc-root \\server\\share` | Override if mapped-drive → UNC fails |
| `--import-excel "..\\Lan_Search_Tool.xlsm"` | Workbook for sheet import or AccDB relative `\DB\` path |

### Crawl + AccDB / Excel

```bat
python -m crawler "O:\Some Folder" -w 16 --import-excel "..\Lan_Search_Tool.xlsm" --target AccDB
```

### Portable zip

On a machine that can `pip download` (JFrog):

```bat
pack_portable.cmd
```

Creates `dist/Lan_Search_Tool.zip`. Colleague unzips, double-clicks `run_crawl.cmd`. Packages install **offline** from `vendor/wheels` into `.portable_venv` (no PyPI). After the crawl they can delete that venv.

The packer downloads wheels for Python **3.10–3.14** (32-bit and 64-bit Windows). If install fails, the cmd window prints this PC’s Python vs the packed `cp###` wheel names.

## Output columns (same as VBA `tblFiles`)

| Column | Meaning |
|--------|---------|
| FilePath | UNC path |
| FileDate | `yyyy-mm-dd` (**files: last modified**; folders: created) |
| SizeMB | 2 decimals; min 0.01 |
| EntryType | `FILE` or `FOLDER` |

AccDB also stores `tblFolderMeta` (directory mtime + local bytes) so the next **quick update** can skip unchanged folders.

**FOLDER SizeMB** = recursive disk total of **all** files under that folder (including non-indexed extensions). On a quick skip, folder SizeMB for unchanged trees stays as last fully listed.

## Filters (parity with VBA)

- Allowed extensions: documents, images, parent archives, video (see `filters.py`)
- Skips junk folders (`.git`, `$Recycle.Bin`, …) and junk files (`Thumbs.db`, `~$*`, …)
- Skips split archive volumes (`.r01`, `.z01`, `.part2.rar`, …)

## Layout

```text
src/crawler/
  __main__.py         CLI
  crawl.py            Thread-pool crawler + AccDB write
  crawl_gui.py        Setup + live progress window
  index_state.py      Previous-index snapshots for quick update
  filters.py          Allowlist / junk / split volumes
  paths.py            UNC resolve, size helpers
  import_to_excel.py  Sheet import + clear_onboard_tblfiles
  run_crawl.py        Folder-picker GUI launcher
  summit_lan_scan.py  All mapped-drives GUI launcher

src/pack_portable.py  Zip colleague kit (pack_portable.cmd)
```

## Tips

1. Prefer **UNC** paths when possible (one less mapping hop).
2. If the server feels saturated, **lower** `--workers` (e.g. 8). If it feels idle, try **24–32**.
3. Weekly runs should use **Quick update**. Use **Full recrawl** if you suspect missed files (some NAS boxes do not update folder dates).
4. After AccDB write, reopen the workbook so search resolves the AccDB backend (sheet is cleared).
5. ACE must be installed on machines that crawl to AccDB **and** on machines that search AccDB from Excel.
