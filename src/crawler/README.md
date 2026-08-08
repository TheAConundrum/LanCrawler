# LAN Search Tool — Parallel Python Crawler

Faster LAN indexing than the VBA `FileSystemObject` crawl. Walks folders with a **thread pool** + `os.scandir`, applies the same allowlist / junk / split-archive rules as `modPathUtil`, and writes `tblFiles`-compatible CSV (and optional SQLite).

## Why it’s faster

LAN crawls are dominated by waiting on directory listings. Multiple workers overlap those waits. VBA cannot do this cleanly inside Excel.

## Requirements

- Python 3.10+ on Windows
- Stdlib only for crawl/CSV/SQLite
- Optional Excel import: `pip install -r requirements-crawler.txt` (xlwings)

## Quick start (easiest in Cursor)

**Run Task:** `Crawl LAN folder (picker)`  
or **Run and Debug:** `Crawl LAN folder (picker)`  
or right-click / Run: `src/crawler/run_crawl.py`

A Windows folder dialog opens (check the taskbar if it’s behind Cursor). After the crawl it **auto-imports** into `Blank_LAN_Crawler Tool.xlsm` (no prompts): events/screen off → write → save → close → reopen so `Workbook_Open` warms the search cache. CSV still goes to `crawl_output\`.

Direct file run also works on `crawl.py` / `run_crawl.py` (same flow).

UNC also works:

```bat
python -m crawler "\\fileserver\share$\Some Folder" -w 24
```

Outputs default to `crawl_output\<timestamp>_tblFiles.csv`.

### Common flags

| Flag | Meaning |
|------|---------|
| `-w` / `--workers` | Parallel folder workers (default **16**; try 8–32) |
| `-o path.csv` | CSV output path |
| `--sqlite path.sqlite` | Also write SQLite `tblFiles` |
| `--unc-root \\server\share` | Override if mapped-drive → UNC fails |
| `--import-excel "..\Blank_LAN_Crawler Tool.xlsm"` | Push CSV into `Database!tblFiles` after crawl |
| `--mode ReplaceRoot` | Import mode: `ReplaceRoot` (default), `Append`, `RebuildAll` |

### Crawl + import in one step

```bat
python -m crawler "O:\Some Folder" -w 16 --import-excel "..\Blank_LAN_Crawler Tool.xlsm" --mode ReplaceRoot
```

Excel import always: close target if open → open with **events/screen off** → write `tblFiles` + Ingestion summary → **enforce UI lock** (keywords+MB only; Database very hidden) → save → close → **reopen with events on** (cache warm). Excel stays open at the end. Does **not** touch the VBA project password (set that once in the VBE as admin).

## Output columns (same as VBA `tblFiles`)

| Column | Meaning |
|--------|---------|
| FilePath | UNC path |
| FileDate | `yyyy-mm-dd` (Windows creation time when available) |
| SizeMB | 2 decimals; min 0.01 |
| EntryType | `FILE` or `FOLDER` |

**FOLDER SizeMB** = recursive disk total of **all** files under that folder (including non-indexed extensions), matching the VBA crawler.

## Filters (parity with VBA)

- Allowed extensions: documents, images, parent archives, video (see `filters.py`)
- Skips junk folders (`.git`, `$Recycle.Bin`, …) and junk files (`Thumbs.db`, `~$*`, …)
- Skips split archive volumes (`.r01`, `.z01`, `.part2.rar`, …)

## Layout

```text
src/crawler/
  __main__.py      CLI
  crawl.py         Thread-pool crawler
  filters.py       Allowlist / junk / split volumes
  paths.py         UNC resolve, size helpers
  import_to_excel.py
```

## Tips

1. Prefer **UNC** paths when possible (one less mapping hop).
2. If the server feels saturated, **lower** `--workers` (e.g. 8). If it feels idle, try **24–32**.
3. After import, wait ~2s for the deferred index cache warm (Immediate Window shows `INDEX CACHE` lines).
4. Large ReplaceRoot imports still rewrite the Excel table once at the end — the slow part is the crawl, which is now parallel.
