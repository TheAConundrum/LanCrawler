"""
CLI: parallel LAN crawler for LAN Search Tool.

Examples:
  python -m crawler "O:\\Some Folder" --workers 16
  python -m crawler "\\\\server\\share\\folder" -o out\\index.csv --sqlite out\\index.sqlite
  python -m crawler "O:\\folder" --import-excel "..\\Blank_LAN_Crawler Tool.xlsm" --mode ReplaceRoot
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

# Allow `python -m crawler` from src/ or project root
_SRC = Path(__file__).resolve().parent.parent
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from crawler.crawl import crawl_parallel, write_csv, write_sqlite
from crawler.paths import build_drive_map, bytes_to_size_mb, to_unc_path


def _default_out_dir() -> Path:
    return Path(__file__).resolve().parent.parent.parent / "crawl_output"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Parallel LAN crawler for LAN Search Tool (writes tblFiles-compatible CSV/SQLite)."
    )
    parser.add_argument("root", help="Folder to crawl (mapped drive or UNC)")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="CSV output path (default: crawl_output/<timestamp>_tblFiles.csv)",
    )
    parser.add_argument(
        "--sqlite",
        type=Path,
        default=None,
        help="Optional SQLite output path",
    )
    parser.add_argument(
        "-w",
        "--workers",
        type=int,
        default=16,
        help="Parallel folder workers (default 16). Try 8–32 on LAN.",
    )
    parser.add_argument(
        "--unc-root",
        default="",
        help="UNC root override if drive mapping cannot be resolved",
    )
    parser.add_argument(
        "--progress-every",
        type=int,
        default=50,
        help="Print progress every N folders (default 50)",
    )
    parser.add_argument(
        "--import-excel",
        type=Path,
        default=None,
        help="After crawl, import CSV into this .xlsm (requires xlwings)",
    )
    parser.add_argument(
        "--mode",
        choices=("ReplaceRoot", "Append", "RebuildAll"),
        default="ReplaceRoot",
        help="Excel import mode when --import-excel is set (default ReplaceRoot)",
    )
    args = parser.parse_args(argv)

    out_dir = _default_out_dir()
    stamp = time.strftime("%Y%m%d_%H%M%S")
    csv_path = args.output or (out_dir / f"{stamp}_tblFiles.csv")
    sqlite_path = args.sqlite

    print(f"Root:     {args.root}")
    print(f"Workers:  {args.workers}")
    print(f"CSV out:  {csv_path}")
    if sqlite_path:
        print(f"SQLite:   {sqlite_path}")

    drive_map = build_drive_map([args.root])
    unc = to_unc_path(args.root, drive_map, args.unc_root)
    print(f"UNC:      {unc}")
    if drive_map:
        print(f"Drives:   {drive_map}")

    def on_progress(msg: str) -> None:
        print(msg, flush=True)

    t0 = time.time()
    rows, stats = crawl_parallel(
        args.root,
        workers=args.workers,
        drive_map=drive_map,
        unc_root_override=args.unc_root,
        progress_every=max(1, args.progress_every),
        on_progress=on_progress,
    )

    write_csv(rows, csv_path)
    print(f"Wrote CSV {csv_path} ({len(rows):,} rows) in {time.time() - t0:.1f}s")

    if sqlite_path:
        write_sqlite(rows, sqlite_path)
        print(f"Wrote SQLite {sqlite_path}")

    print(
        f"Stats: files={stats.files_indexed:,} folders={stats.folders_indexed:,} "
        f"errors={stats.errors:,} disk~{bytes_to_size_mb(stats.bytes_all_files):,.2f} MB "
        f"crawl={stats.elapsed:.1f}s"
    )

    if args.import_excel:
        from crawler.import_to_excel import import_csv_to_workbook

        print(f"Importing into {args.import_excel} mode={args.mode} ...")
        import_csv_to_workbook(
            workbook=args.import_excel,
            csv_path=csv_path,
            mode=args.mode,
            unc_root=unc,
        )
        print("Excel import complete.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
