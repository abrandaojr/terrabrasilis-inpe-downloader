# Changelog

All notable changes to this project are documented here.

## Unreleased

- Ensure `vector-to-geoparquet` is upgraded to at least `1.4.1` before
  conversion, preventing stale `proj.db` path conflicts between pyproj and
  GDAL/pyogrio/rasterio.

## 1.2.0 - 2026-06-11

- Reorganized the GitHub repository into `scripts/`, `prodes_pipeline/`, and `docs/`.
- Added `run_pipeline.py` as a stable root entry point.
- Converted reusable helpers into the `prodes_pipeline` package.
- Updated setup, README, contribution, and release instructions for the new layout.
- Kept stage scripts directly runnable from a fresh clone without requiring package installation first.

## 1.1.0 - 2026-06-11

- Made the pipeline portable across machines with configurable workspace paths.
- Added automatic creation of ZIP, GeoParquet, table, figure, presentation, and report folders.
- Added `setup_env.py` to create `.venv`, install `uv`, and install project dependencies.
- Added package metadata in `pyproject.toml`.
- Moved generated outputs under `workspace/` by default.
- Removed local editor metadata and generated binary artifacts from version control.
- Expanded repository documentation and GitHub community files.

## 1.0.0

- Initial TerraBrasilis/PRODES download and conversion pipeline.
