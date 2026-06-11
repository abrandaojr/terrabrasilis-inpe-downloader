# PRODES TerraBrasilis Pipeline

Portable Python scripts to download, convert, validate, analyze, and present
INPE PRODES deforestation data from TerraBrasilis.

Author: Amintas Brandao Jr. `<abrandaojr@gmail.com>`  
Affiliation: Imazon - Instituto do Homem e Meio Ambiente da Amazonia

## What This Repository Does

- Downloads public PRODES ZIP archives from TerraBrasilis.
- Converts vector layers to GeoParquet and rasters to Cloud-Optimized GeoTIFF.
- Builds quality reports, catalogs, charts, Excel workbooks, and PowerPoint decks.
- Creates every required local folder automatically.
- Runs on any machine with Python 3.11+ installed.

No machine-specific path is required. By default, generated files are written to
`workspace/` inside the cloned repository. Set `PRODES_HOME` to use another disk
or folder.

## Quick Start

```bash
git clone https://github.com/abrandaojr/terrabrasilis-inpe-downloader.git
cd terrabrasilis-inpe-downloader
python setup_env.py
```

Activate the environment:

```bash
# Windows PowerShell
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate
```

Run the full pipeline:

```bash
python run_pipeline.py
```

Useful run modes:

```bash
python run_pipeline.py -k
python run_pipeline.py --from 2
python run_pipeline.py --steps 1 3
python run_pipeline.py --from 2 -k
```

The individual stage scripts can also be run directly, for example
`python scripts/01_download_zips.py`. They auto-install missing Python packages
into the active interpreter, but `setup_env.py` is the recommended reproducible
setup path.

## Repository Layout

```text
repository/
  run_pipeline.py              # friendly entry point
  setup_env.py                 # creates .venv and installs dependencies
  scripts/                     # runnable pipeline stages
    00_pipeline.py
    01_download_zips.py
    02_convert_to_geoparquet.py
    03_deforestation_chart.py
    04_generate_presentation.py
    05_organize_geoparquet.py
    06_export_tables.py
    07_visual_story_deliverables.py
  prodes_pipeline/             # shared package code
    config.py
    data_quality.py
    pipeline_contracts.py
  docs/
    RELEASES.md
    SECURITY.md
  .github/                     # issue and pull request templates
```

## Portable Workspace

Default layout:

```text
repository/
  scripts/
  prodes_pipeline/
  workspace/
    zip/
    geoparquet/
    tables/
      charts/
    figures/
    presentations/
    reports/
```

Override the workspace:

```bash
# Windows PowerShell
$env:PRODES_HOME = "D:\prodes-workspace"

# macOS / Linux
export PRODES_HOME="/data/prodes-workspace"
```

Optional advanced overrides:

```text
PRODES_ZIP_ROOT
PRODES_GEOPARQUET_DIR
PRODES_TABLES_DIR
PRODES_FIGURES_DIR
PRODES_REPORTS_DIR
PRODES_PRESENTATIONS_DIR
PRODES_EXTRACT_DIR
```

## Pipeline

```text
TerraBrasilis download page
  -> scripts/01_download_zips.py
  -> scripts/02_convert_to_geoparquet.py
  -> scripts/03_deforestation_chart.py
  -> scripts/04_generate_presentation.py
  -> scripts/05_organize_geoparquet.py
  -> scripts/06_export_tables.py
  -> scripts/07_visual_story_deliverables.py
```

### 01 Download ZIPs

Discovers, downloads, resumes, and validates TerraBrasilis ZIP archives.

Key outputs:

- `workspace/zip/YYYY-MM-DD/**/*.zip`
- `workspace/zip/YYYY-MM-DD/terrabrasilis_zips.csv`
- `workspace/zip/YYYY-MM-DD/terrabrasilis_zips.json`
- `workspace/reports/*.json`

### 02 Convert to GeoParquet and COG

Converts vector data to GeoParquet and raster data to Cloud-Optimized GeoTIFF.

Key outputs:

- `workspace/geoparquet/**/*.parquet`
- `workspace/geoparquet/**/*.tif`
- `workspace/reports/02_gpkg_layer_cache.json`

### 03 Chart

Creates a publication-ready annual deforestation chart.

Key output:

- `workspace/figures/amazon_deforestation_norad.png`

### 04 Presentation

Creates a bilingual press briefing PowerPoint.

Key output:

- `workspace/presentations/PRODES_Press_Briefing.pptx`

### 05 Organize GeoParquet

Builds a cleaner cataloged GeoParquet folder structure.

Key output:

- `workspace/geoparquet/_organized/`

### 06 Export Tables

Exports analytical tables and charts.

Key outputs:

- `workspace/tables/PRODES_Analytics_*.xlsx`
- `workspace/tables/PRODES_Analytics_*.pptx`
- `workspace/tables/charts/*.png`

### 07 Visual Story Deliverables

Creates additional didactic PowerPoint and Excel deliverables.

Key outputs:

- `workspace/presentations/PRODES_VISUAL_STORY_*.pptx`
- `workspace/tables/PRODES_VISUAL_STORY_*.xlsx`
- `workspace/figures/prodes_annual_deforestation_*.png`

## Dependencies

Required:

- Python 3.11+
- Internet access for TerraBrasilis downloads and first-time dependency install
- Chrome or Chromium only if the Selenium fallback is needed for JavaScript
  rendering

Recommended setup:

```bash
python setup_env.py
```

Manual setup:

```bash
python -m venv .venv
.venv\Scripts\activate      # Windows
python -m pip install --upgrade pip uv
python -m uv pip install -r requirements.txt
```

On macOS/Linux, activate with `source .venv/bin/activate`.

## Quality and Observability

Each stage writes JSON quality reports to `workspace/reports/`, including:

- input and output inventory
- freshness checks
- row and file counts
- lineage records
- stage timing

Run a syntax check before publishing changes:

```bash
python -m py_compile run_pipeline.py setup_env.py scripts/*.py prodes_pipeline/*.py
```

## Releases and Packages

- Release notes live in `CHANGELOG.md`.
- Release process notes live in `docs/RELEASES.md`.
- Package metadata lives in `pyproject.toml`.
- Generated data products are intentionally not published as GitHub Packages.

## Citation

If you use this repository, cite this project and cite INPE/PRODES according to
the official data provider requirements.

## License

MIT. See `LICENSE`.
