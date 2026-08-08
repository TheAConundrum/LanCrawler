"""
Double-click / right-click Run launcher for the LAN crawler.

Prefer running THIS file if crawl.py does nothing from Cursor.
"""
from __future__ import annotations

import sys
from pathlib import Path

# Ensure src/ is on path, then run interactive crawl
_SRC = Path(__file__).resolve().parent.parent
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from crawler.crawl import run_interactive  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(run_interactive())
