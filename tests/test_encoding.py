from __future__ import annotations

import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
TEXT_GLOBS = ("*.py", "*.md", "*.toml", "*.yml", "*.yaml", ".gitignore")
SKIP_DIRS = {".git", ".venv", "__pycache__", "workspace", "reports", "figures"}
MOJIBAKE_MARKERS = (
    "\u00c3",
    "\u00c2",
    "\u00e2\u20ac",
    "\u00e2\u2020",
    "\u00e2\u20ac\u201c",
    "\u00e2\u20ac\u201d",
    "\u00ce",
    "\u00cf",
    "\u00e2\u2030",
    "\ufffd",
)


def iter_text_files() -> list[Path]:
    files: list[Path] = []
    for pattern in TEXT_GLOBS:
        for path in PROJECT_ROOT.rglob(pattern):
            if any(part in SKIP_DIRS for part in path.parts):
                continue
            files.append(path)
    return sorted(set(files))


class EncodingTests(unittest.TestCase):
    def test_source_text_has_no_common_mojibake_markers(self) -> None:
        offenders: list[str] = []
        for path in iter_text_files():
            text = path.read_text(encoding="utf-8")
            markers = [marker for marker in MOJIBAKE_MARKERS if marker in text]
            if markers:
                rel = path.relative_to(PROJECT_ROOT)
                offenders.append(f"{rel}: {', '.join(markers)}")

        self.assertEqual(offenders, [])


if __name__ == "__main__":
    unittest.main()
