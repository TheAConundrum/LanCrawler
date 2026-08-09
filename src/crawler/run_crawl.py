"""
Double-click / right-click Run launcher for the LAN crawler.

Prefer running THIS file if crawl.py does nothing from Cursor.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

# Ensure src/ is on path, then run interactive crawl
_SRC = Path(__file__).resolve().parent.parent
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from crawler.crawl import run_interactive  # noqa: E402


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="LAN crawler (GUI: folder + workbook + AccDB target)")
    parser.add_argument(
        "--workbook",
        default=None,
        help="Target .xlsm (skips GUI workbook field if set with --no-gui)",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=16,
        help="Parallel folder workers (default 16)",
    )
    parser.add_argument(
        "--target",
        choices=("Auto", "Workbook", "AccDB"),
        default=None,
        help="Index target mode (default Auto via GUI)",
    )
    parser.add_argument(
        "--no-gui",
        action="store_true",
        help="Use folder/workbook pickers instead of the crawl GUI",
    )
    parser.add_argument(
        "--no-excel",
        action="store_true",
        help="Write CSV only; do not import into Excel / AccDB",
    )
    args = parser.parse_args(argv)
    return run_interactive(
        workers=args.workers,
        auto_import_excel=not args.no_excel,
        workbook=args.workbook,
        target_mode=args.target,
        use_gui=not args.no_gui,
    )


if __name__ == "__main__":
    raise SystemExit(main())
