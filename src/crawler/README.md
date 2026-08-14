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
or right-click / Run: `src/crawler/run_crawl.py`

A small GUI asks for:

1. Folder / drive to crawl  
2. Target workbook (defaults to `AccDB-Blank_LAN_Crawler Tool.xlsm`)  
3. Index target: **Auto** | **Workbook** | **AccDB**

**Auto / Workbook** under **750,000** rows → import into `Database!tblFiles`.  
**AccDB**, or any mode above **750,000** rows → write `{workbook}\DB\LAN_Search_Index.accdb` (`tblFiles` + `tblIngested`). Each AccDB crawl **ReplaceRoot**-merges that folder into the AccDB (keeps other roots’ files and scan history). On workbook open / cache warm (and when AccDB is deleted), VBA restores **Ingestion!tblIngested** from AccDB only — or clears it if the AccDB file is gone. CSV still goes to `crawl_output\`.

Keep `Blank_LAN_Crawler Tool.xlsm` as the stable non-AccDB template; AccDB development uses `AccDB-Blank_LAN_Crawler Tool.xlsm`.

UNC also works:

```bat
python -m crawler "\\fileserver\share$\Some Folder" -w 24
```

### Common flags

| Flag | Meaning |
|------|---------|
| `-w` / `--workers` | Parallel folder workers (default **16**; try 8–32) |
| `-o path.csv` | CSV output path |
| `--accdb path.accdb` | Write AccDB `tblFiles` |
| `--target Auto\|Workbook\|AccDB` | Exclusive index mode with `--import-excel` |
| `--unc-root \\server\share` | Override if mapped-drive → UNC fails |
| `--import-excel "..\AccDB-Blank_LAN_Crawler Tool.xlsm"` | Workbook for sheet import or AccDB relative `\DB\` path |

### Crawl + AccDB / Excel

```bat
python -m crawler "O:\Some Folder" -w 16 --import-excel "..\AccDB-Blank_LAN_Crawler Tool.xlsm" --target Auto
```

## Output columns (same as VBA `tblFiles`)

| Column | Meaning |
|--------|---------|
| FilePath | UNC path |
| FileDate | `yyyy-mm-dd` (Windows creation time when available) |
| SizeMB | 2 decimals; min 0.01 |
| EntryType | `FILE` or `FOLDER` |

**FOLDER SizeMB** = recursive disk total of **all** files under that folder (including non-indexed extensions).

## Filters (parity with VBA)

- Allowed extensions: documents, images, parent archives, video (see `filters.py`)
- Skips junk folders (`.git`, `$Recycle.Bin`, …) and junk files (`Thumbs.db`, `~$*`, …)
- Skips split archive volumes (`.r01`, `.z01`, `.part2.rar`, …)

## Layout

```text
src/crawler/
  __main__.py         CLI
  crawl.py            Thread-pool crawler + write_accdb
  crawl_gui.py        Small target GUI
  filters.py          Allowlist / junk / split volumes
  paths.py            UNC resolve, size helpers
  import_to_excel.py  Sheet import + clear_onboard_tblfiles
  run_crawl.py        Launcher
```

## Tips

1. Prefer **UNC** paths when possible (one less mapping hop).
2. If the server feels saturated, **lower** `--workers` (e.g. 8). If it feels idle, try **24–32**.
3. After AccDB write, reopen the workbook so search resolves the AccDB backend (sheet is cleared).
4. ACE must be installed on machines that crawl to AccDB **and** on machines that search AccDB from Excel.
