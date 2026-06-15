from __future__ import annotations

import unittest
from pathlib import Path

from prodes_pipeline.config import ANALYSIS_BASE_YEAR, DATA_YEAR_MAX, DATA_YEAR_MIN


PROJECT_ROOT = Path(__file__).resolve().parents[1]


class AnalysisPolicyTests(unittest.TestCase):
    def test_base_year_policy_is_centralized(self) -> None:
        self.assertEqual(ANALYSIS_BASE_YEAR, 2008)
        self.assertLess(DATA_YEAR_MIN, ANALYSIS_BASE_YEAR)
        self.assertGreater(DATA_YEAR_MAX, ANALYSIS_BASE_YEAR)

    def test_presentation_stats_use_2008_baseline(self) -> None:
        source = (PROJECT_ROOT / "scripts" / "04_generate_presentation.py").read_text(
            encoding="utf-8"
        )

        self.assertIn("analytical_amazon = {y: v for y, v in amazon.items() if y >= base_year}", source)
        self.assertIn('s["baseline_year"]', source)
        self.assertIn('s["baseline_km2"]', source)
        self.assertIn('(s["baseline_km2"] - s["current_km2"])', source)

    def test_export_analytics_filter_to_base_year_after_context_series(self) -> None:
        source = (PROJECT_ROOT / "scripts" / "06_export_tables.py").read_text(
            encoding="utf-8"
        )

        self.assertIn("def _rows_from_year", source)
        self.assertIn("p1_analysis = _rows_from_year(p1)", source)
        self.assertIn("p2 = p2_cumulative(p1_analysis)", source)
        self.assertIn("BETWEEN {ANALYSIS_BASE_YEAR} AND {DATA_YEAR_MAX}", source)
        self.assertIn('normalized = part.replace("_", " ")', source)
        self.assertIn("biome = _find_biome(parts)", source)

    def test_secondary_vegetation_aliases_are_recognized_and_excluded(self) -> None:
        presentation = (
            PROJECT_ROOT / "scripts" / "04_generate_presentation.py"
        ).read_text(encoding="utf-8")
        exports = (PROJECT_ROOT / "scripts" / "06_export_tables.py").read_text(
            encoding="utf-8"
        )

        for source in (presentation, exports):
            self.assertIn("floresta secundaria", source)
            self.assertIn("vegetacao secundaria", source)
            self.assertIn("secondary vegetation", source)
            self.assertIn('"vs"', source)

        self.assertIn("if _is_secondary_vegetation_path(parts):", presentation)
        self.assertIn("if _is_vs_path(parts):", exports)


if __name__ == "__main__":
    unittest.main()
