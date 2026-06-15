from __future__ import annotations

import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]


class OutputPathTests(unittest.TestCase):
    def test_analytics_charts_are_saved_to_figures_dir(self) -> None:
        source = (PROJECT_ROOT / "scripts" / "06_export_tables.py").read_text(
            encoding="utf-8"
        )
        readme = (PROJECT_ROOT / "README.md").read_text(encoding="utf-8")

        self.assertIn("FIGURES_DIR", source)
        self.assertIn("CHART_DIR = FIGURES_DIR", source)
        self.assertNotIn('TABLES_DIR / "charts"', source)
        self.assertIn("workspace/figures/*.png", readme)
        self.assertNotIn("workspace/tables/charts/*.png", readme)


if __name__ == "__main__":
    unittest.main()
