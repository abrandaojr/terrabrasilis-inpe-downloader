from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from prodes_pipeline.data_quality import LineageRecord, atomic_write_json


class AtomicWriteJsonTests(unittest.TestCase):
    def test_writes_lineage_record_as_json_object(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "report.json"
            atomic_write_json(
                path,
                {
                    "lineage": LineageRecord(
                        stage_name="02_convert_to_geoparquet",
                        upstream_sources=["input.zip"],
                        transformation="convert",
                        downstream_outputs=["workspace/geoparquet"],
                        contracts=["zip_archive", "geoparquet"],
                    )
                },
            )

            payload = json.loads(path.read_text(encoding="utf-8"))

        self.assertEqual(payload["lineage"]["stage_name"], "02_convert_to_geoparquet")
        self.assertEqual(payload["lineage"]["upstream_sources"], ["input.zip"])


if __name__ == "__main__":
    unittest.main()
