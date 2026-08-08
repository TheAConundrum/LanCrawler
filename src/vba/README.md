# LAN Search Tool — VBA import guide

## Preferred: VS Code task (silent inject)

1. Trust access to the VBA project object model enabled in Excel  
   (File → Options → Trust Center → Trust Center Settings → Macro Settings)
2. Run task **Import VBA code to Excel**  
   - Closes `Blank_LAN_Crawler Tool.xlsm` if open  
   - Injects everything from `src/vba/`  
   - Removes old `Module1`–`Module4` if present  
   - Leaves Excel open with the updated workbook  

   Companion task: **Export from Excel to Disk** (VBA → `src/vba`)  

Or from a terminal:

```text
python src\inject_vba.py "Blank_LAN_Crawler Tool.xlsm"
```

Manual import below is only a fallback.

## Sheet rename map

| Old | New |
|-----|-----|
| HomePage | Dashboard |
| Feed The Beast | Ingestion |
| Data | Database |

Run `RunSetupWorkbook` once after import to rename sheets (if still old names), apply labels, and set dropdown validation.

## Modules to import

| File | Type |
|------|------|
| `modPathUtil.bas` | Standard |
| `modEntryPoints.bas` | Standard |
| `clsSheetMap.cls` | Class |
| `clsExcelAppState.cls` | Class |
| `clsFileIndexer.cls` | Class |
| `clsFileSearch.cls` | Class |

Remove old `Module1`–`Module4` (and `Macro9`) after import.

## Button rebinds (only these two are public)

| Old macro | New macro |
|-----------|-----------|
| `ListFiles` | `RunIndexFromIngestion` |
| `GetSearchResults` | `RunSearchFromDashboard` |

Class methods and `modPathUtil` helpers are hidden from the Assign Macro list (`Option Private Module`).  
If you still see `ListFiles` / `GetSearchResults` / many other names, **delete the old Module1–Module4** from the VBA project.

## Cell layout

### Dashboard (user owns rows 1–6 chrome)

| Row | Purpose |
|-----|---------|
| 1 | Instructions (yours) |
| 2 | Criteria headers |
| 3–5 | Criteria inputs |
| 3–6 | Options / stats chrome (right side; yours) |
| 8 | Divider (yours — macros never write here) |
| 9 | Results headers (user-owned — macros never change text/formatting) |
| 10+ | Results data only (macro-drawn stripes; not an Excel ListObject) |

Criteria (`A2:F5`): Order | Key Word | Operator | Key Word | Size | Size (MB)  
Flexible match: label **`H3`**, value **`H4`**  
Search folders: label **`I3:J3`**, value **`I4`** = `Yes` (default) / `No` / `Only`  
- **Yes** — file names + folder names (folders collapsed with counts/dates)  
- **No** — file names only  
- **Only** — folder names only  

Drive filter: value **`H5`** (merged `H5:J5` ok) — dropdown populated when the index cache loads. Default **`All`**, or a friendly volume name like `SHARE_Public`. Restricts search to that UNC share.

Search stats: macro writes **`M3`** file count, **`M4`** folder count, **`M5`** total Size MB (merged ranges `M3:N3` / `M4:N4` / `M5:N5` — value goes in the top-left cell).

Results headers (row 7) / data from row 8: **A:K** File (blue; **double-click** opens containing folder) | **L** File Type | **M** Size MB | **N** Date Created | **O** Drive Address (friendly label, e.g. `SHARE_Public`) | **P:T** Folder Path (plain text, relative, leading `\`)

Hidden column **U** stores the UNC folder target for double-click (not shown). Search does **not** create Excel hyperlink objects. **Double-click a File name** to open its folder.

**UI lock:** only free-text criteria are editable — Key Word columns **B** and **D**, and Size MB **F** (rows 3–5). Operators, Flexible match, Search folders, Drive filter, results, stats, and Ingestion are locked (`EnableSelection = xlNoRestrictions` avoids Excel’s protect popup). **Database** is sheet-protected and **`xlSheetVeryHidden`** (not in Unhide). VBA search/ingest and the Python crawler import unprotect/unhide as needed, then re-lock.

**VBA project password:** set once as admin in VBE → Tools → VBAProject Properties → Protection. Python does **not** set or bypass it. The crawler can still write sheet data (Database/Ingestion) while the project is locked; **Import VBA to Excel** cannot modify modules until you unlock the project for that edit session.

`Database!tblFiles` **FilePath** and Ingestion **RootPath** stay as full UNC — needed for open-folder targets, ReplaceRoot matching, and crawl import. Friendly names are display-only (drive dropdown + Drive Address column).

### Ingestion

- **No path text box** — `RunIndexFromIngestion` opens a folder picker
- `B2` — UncRoot override (if mapped-drive → UNC resolve fails)
- `B4` — Mode: ReplaceRoot / Append / RebuildAll (`RebuildAll` uses Links already Eaten, no picker)
- **Summary** (compact): row 1 header (yours); **D2:D5** labels (user-owned) | **E2:E5** values refreshed on ingest — Scans, Files, Folders, Total Size MB. Macro writes values only (no label/format changes).
- `tblIngested` (from **A6**): **RootPath** (UNC) | **DateIngested** | **FileCount** | **FolderCount** | **TotalSizeMB**

### Database

- `tblFiles`: **FilePath** | **FileDate** | **SizeMB** | **EntryType** (`FILE` / `FOLDER`)
  - FILE SizeMB = that file. FOLDER SizeMB = recursive disk total under the folder (all files seen during crawl, not only allowlisted).

## Index filter (docs + images + parent archives)

Crawl indexes only these extensions (see `modPathUtil.ALLOWED_INDEX_EXTS`):

- Documents: `pdf doc docx xls xlsx xlsm xlsb ppt pptx txt rtf csv odt ods odp msg eml`
- Images: `jpg jpeg png gif tif tiff bmp webp`
- Archives (parent only): `zip 7z rar tar gz tgz bz2 xz cab`
- Video: `mp4 mov avi mkv wmv webm m4v mpg mpeg flv 3gp ts mts m2ts`

Multi-volume siblings are skipped when the parent is enough, e.g. keep `tmllist.rar` / `name.part1.rar`, skip `tmllist.r01`…`r100`, `name.z01`, `name.7z.001`, `name.part2.rar`.

Also skips junk folders (`.git`, `__pycache__`, `$Recycle.Bin`, …) and junk files (`Thumbs.db`, `~$*`, `.tmp`, …).

## Parallel Python crawler (recommended for large LAN roots)

See [`src/crawler/README.md`](../crawler/README.md). Example:

```bat
cd /d "D:\path\to\LAN Search Tool\src"
python -m crawler "O:\Some Folder" -w 16 --import-excel "..\Blank_LAN_Crawler Tool.xlsm"
```

Uses a thread pool + `os.scandir` (same filters as VBA). Much faster on LAN than the Excel crawler; VBA search/UI stays the same after CSV import into `tblFiles`.

## Search behavior

- **Files:** keyword match against the **file name** only.
- **Folders** (`I4`: Yes / No / Only): match indexed **folder** rows by name; one result row with recursive indexed file/folder counts, Date Created from ingest, and **Size MB** = recursive disk total of all files under that folder (including extensions not indexed), written at crawl time.
- Flexible match (`H4`) applies to those name tests.
- Drive filter (`H5`) limits hits to one share when not `All`.

## Search performance

`modIndexCache` keeps `tblFiles` in memory for the workbook session, plus parallel search arrays (path, file/folder name, flexible-normalized name, size, date, parent links, descendant file/folder counts) built once on warm load / ingest. Search scans with plain `InStr`; folder result rows are O(hits) lookups against those precomputed counts.

`Workbook_Open` schedules a deferred warm load (~2s later via `Application.OnTime`) so Excel opens immediately, then builds the results display grid if needed. Immediate Window (`Ctrl+G`) prints `SEARCH TIMING` / `INDEX CACHE` diagnostics.

Results are written into a **prebuilt 5000-row grid** (File `A:K` and Folder `P:T` merges + grey/white striping). That grid is built once on cache warm; each search only clears values and batch-writes into the top-left merge cells (plus hidden UNC targets in column U). **Double-click a File name** to open its folder in Explorer.

## Notes

- Re-ingest the same root to refresh counts/sizes (ReplaceRoot).
- Size filter is optional; leave Size / Size (MB) blank to ignore size.
- Flexible match strips ` .,;-_` etc.; keeps `\` and `/`.
- On ingest failure: read the MsgBox **Stage** / **Error #**, and open VBA Immediate Window (**Ctrl+G**) for the full log.
- Name the Links already Eaten Excel Table `tblIngested` (or keep a header cell `RootPath`) — the macro will not overwrite your Ingestion layout labels.
