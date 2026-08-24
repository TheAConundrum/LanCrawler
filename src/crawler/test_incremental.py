"""Incremental crawl: resaves and nested new files."""
from __future__ import annotations

import sys
import tempfile
import time
import unittest
from pathlib import Path

_SRC = Path(__file__).resolve().parent.parent
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from crawler.crawl import crawl_parallel  # noqa: E402
from crawler.index_state import (  # noqa: E402
    FolderMetaRow,
    build_previous_index,
    folder_mtime_unchanged,
)


def _write(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _previous_from_rows(rows):
    file_rows = [(r.path, r.file_date, r.size_mb, r.entry_type) for r in rows]
    meta = [
        FolderMetaRow(
            folder_path=r.path,
            dir_mtime=r.mtime_ts,
            local_bytes=r.size_bytes,
            created_date=r.file_date,
        )
        for r in rows
        if r.entry_type == "FOLDER"
    ]
    return build_previous_index(file_rows=file_rows, meta_rows=meta)


class IncrementalCrawlTests(unittest.TestCase):
    def test_folder_mtime_key(self) -> None:
        self.assertTrue(folder_mtime_unchanged(100.4, 100.9))
        self.assertFalse(folder_mtime_unchanged(101.1, 100.9))
        self.assertFalse(folder_mtime_unchanged(50.0, 0.0))

    def test_resave_and_nested_new_file(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            keep = root / "keep"
            child = root / "child"
            _write(keep / "alpha.txt", "one")
            _write(child / "beta.txt", "two")

            rows1, stats1 = crawl_parallel(str(root), workers=2, progress_every=1)
            self.assertGreaterEqual(stats1.folders_scanned, 3)
            self.assertEqual(stats1.folders_quick, 0)
            names1 = {Path(r.path).name.lower() for r in rows1 if r.entry_type == "FILE"}
            self.assertIn("alpha.txt", names1)
            self.assertIn("beta.txt", names1)

            prev = _previous_from_rows(rows1)
            self.assertTrue(prev.has_meta)

            # Resave in an existing folder. Sleep so mtime date can change.
            time.sleep(1.1)
            _write(keep / "alpha.txt", "one-updated-and-longer")
            _write(child / "gamma.txt", "brand new")

            rows2, stats2 = crawl_parallel(
                str(root),
                workers=2,
                progress_every=1,
                incremental=True,
                previous=prev,
            )
            files2 = [r for r in rows2 if r.entry_type == "FILE"]
            by_name = {Path(r.path).name.lower(): r for r in files2}
            self.assertIn("gamma.txt", by_name)
            self.assertIn("alpha.txt", by_name)
            self.assertEqual(by_name["alpha.txt"].file_date, time.strftime("%Y-%m-%d"))
            self.assertGreater(by_name["alpha.txt"].size_bytes, 3)
            # Unchanged folders should be quick-visited, not all full-listed.
            self.assertGreater(stats2.folders_quick, 0)
            self.assertGreater(stats2.folders_scanned, 0)


if __name__ == "__main__":
    unittest.main()
