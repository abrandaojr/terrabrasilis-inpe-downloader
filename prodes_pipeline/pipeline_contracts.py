from __future__ import annotations

"""Data contracts for the PRODES pipeline stage boundaries.

These contracts encode the assumptions each downstream stage makes about data
produced upstream. They are intentionally lightweight and use only the local
`data_quality` helpers so the pipeline remains dependency-compatible.
"""

from prodes_pipeline.data_quality import DataContract, FreshnessPolicy


ZIP_INVENTORY_CONTRACT = DataContract(
    name="terrabrasilis_zip_inventory",
    producer="01_download_zips.py",
    consumers=("02_convert_to_geoparquet.py",),
    required_columns=("url", "filename", "filename_local", "biome", "category"),
    freshness=FreshnessPolicy(max_age_hours=24 * 14, severity="warn"),
    min_rows=50,
    notes="Remote TerraBrasilis inventory after skip_files exclusions.",
)

ZIP_ARCHIVE_CONTRACT = DataContract(
    name="local_zip_archives",
    producer="01_download_zips.py",
    consumers=("02_convert_to_geoparquet.py",),
    freshness=FreshnessPolicy(max_age_hours=24 * 90, severity="warn"),
    min_rows=50,
    notes="Non-empty, valid ZIP archives under the configured PRODES ZIP root.",
)

GEOPARQUET_CONTRACT = DataContract(
    name="prodes_geoparquet_outputs",
    producer="02_convert_to_geoparquet.py",
    consumers=(
        "04_generate_presentation.py",
        "05_organize_geoparquet.py",
        "06_export_tables.py",
    ),
    optional_columns=(
        "year",
        "ano",
        "state",
        "estado",
        "uf",
        "municipality",
        "municipio",
        "município",
        "classname",
        "class_name",
        "area_km2",
        "area",
        "geometry",
    ),
    numeric_columns=("year", "ano", "area_km2", "area"),
    categorical_columns=("state", "estado", "uf", "classname", "class_name"),
    freshness=FreshnessPolicy(max_age_hours=24 * 120, severity="warn"),
    min_rows=1,
    max_null_rate=0.40,
    notes="Vector GeoParquet outputs and COG rasters generated from PRODES ZIPs.",
)

ANALYTICS_EXPORT_CONTRACT = DataContract(
    name="prodes_analytics_exports",
    producer="06_export_tables.py",
    consumers=("analysts", "presentations", "external_reporting"),
    freshness=FreshnessPolicy(max_age_hours=24 * 30, severity="warn"),
    min_rows=1,
    notes="Excel and PPTX analytic artifacts written under the configured tables workspace.",
)
