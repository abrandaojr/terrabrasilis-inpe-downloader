from __future__ import annotations

import importlib.util
import tempfile
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_PATH = PROJECT_ROOT / "scripts" / "05_organize_geoparquet.py"


def load_organizer_module():
    spec = importlib.util.spec_from_file_location("organize_geoparquet", SCRIPT_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load {SCRIPT_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class OrganizeGeoParquetTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.module = load_organizer_module()

    def test_discover_files_includes_existing_organized_tree(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            existing = (
                root
                / "_organized"
                / "Amazon_Biome"
                / "deforestation"
                / "yearly_deforestation.parquet"
            )
            existing.parent.mkdir(parents=True)
            existing.write_bytes(b"placeholder")

            discovered = self.module.discover_files(root)

        self.assertEqual(
            discovered,
            {("Amazon Biome", "deforestation", "yearly_deforestation.parquet"): existing},
        )

    def test_classifies_vs_and_floresta_secundaria_as_secondary_vegetation(self) -> None:
        cases = [
            ("Amazon_Biome", "VS_2020", "VS_Amazon_Biome.parquet"),
            ("Legal_Amazon", "floresta_secundaria", "floresta_secundaria_2020.parquet"),
            ("Cerrado", "vegetação secundária", "layer.parquet"),
            ("Pampa", "secondary-vegetation", "secondary_forest.parquet"),
        ]

        for parts in cases:
            with self.subTest(parts=parts):
                self.assertEqual(
                    self.module._classify(parts),
                    "secondary_vegetation",
                )


if __name__ == "__main__":
    unittest.main()
