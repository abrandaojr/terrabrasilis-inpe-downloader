# PRODES Data Pipeline

A collection of Python scripts to download, convert, and visualize deforestation data from INPE's [TerraBrasilis](https://terrabrasilis.dpi.inpe.br/en/download-files/) platform (PRODES).

**Author:** Amintas Brandão Jr. \<abrandaojr@gmail.com\>  
**Affiliation:** Imazon — Instituto do Homem e Meio Ambiente da Amazônia

---

## Pipeline Overview

```
TerraBrasilis website
        │
        ▼
01_download_zips.py           ← scrape + download all .zip files
        │
        ▼  (ZIP archives)
02_convert_to_geoparquet.py   ← convert vectors → GeoParquet  (DuckDB-optimized)
        or
03_convert_tiled_vectors_rasters.py  ← convert vectors + rasters → tiled GeoParquet / COG GeoTIFF
        │
        ▼  (analysis-ready files)
04_deforestation_chart.py     ← generate annual deforestation rate chart
```

---

## Scripts

### `01_download_zips.py`

Discovers, downloads, and validates all `.zip` files from TerraBrasilis.

**Features:**
- Static HTML scraping (falls back to Selenium if JS rendering is required)
- Inventory table showing already-downloaded vs. pending files
- Resume support — interrupted downloads continue from where they stopped
- ZIP integrity check + automatic repair loop
- Saves a `validation_report.json` and a CSV/JSON index of all files found

**Configuration** (top of file):
| Variable | Default | Description |
|---|---|---|
| `BASE_URL` | TerraBrasilis download page | Source URL to scrape |
| `ROOT_FOLDER` | `C:\Amintas\Prodes\zip` | Root storage directory |
| `DEST_FOLDER` | `ROOT_FOLDER\<today>` | Download destination |

**Usage:**
```bash
python 01_download_zips.py
```

---

### `02_convert_to_geoparquet.py`

Batch-converts PRODES ZIP archives to DuckDB-optimized GeoParquet files using an external converter fetched from GitHub.

**Features:**
- Incremental — skips already-converted files
- Parallel conversion via thread pool
- Supports `.gpkg` (all layers) and `.shp` (orphan shapefiles)
- Generates a timestamped JSON report per run
- `--list` flag to inspect existing GeoParquets and their CRS

**Configuration** (the `CONFIG` dict at the top of the file):
| Key | Default | Description |
|---|---|---|
| `source_dir` | `C:\Amintas\Prodes\zip\2026-05-07` | Directory with input ZIPs |
| `dest_dir` | `C:\Amintas\Prodes\geoparquet` | Output directory |
| `n_workers` | `8` | Parallel conversion threads |
| `compression` | `zstd` | Parquet compression codec |
| `tile_size_m` | `25000` | Hilbert sort tile size (metres) |

**Usage:**
```bash
python 02_convert_to_geoparquet.py          # convert
python 02_convert_to_geoparquet.py --list   # list existing GeoParquets
```

---

### `03_convert_tiled_vectors_rasters.py`

Alternative converter with full raster support, spatial tiling, and CRS reprojection.

**Features:**
- Reprojects all data to **ESRI:102033** (South America Equidistant Conic)
- Splits vectors into 25 × 25 km tiles aligned to a global grid; adds `tile_id`, `tile_col`, `tile_row` attributes
- Converts rasters to Cloud-Optimized GeoTIFF (COG) tiles with internal overviews
- Geometry repair (`buffer(0)` + `make_valid`) and polygon-type filtering
- Parallel processing via `ThreadPoolExecutor`
- JSON processing report

**Configuration** (constants at the top of the file):
| Variable | Default | Description |
|---|---|---|
| `PASTA_ZIPS` | `C:\Amintas\Prodes\zip\<today>` | Input ZIP directory |
| `PASTA_VET` | `PASTA_ZIPS\geoparquet` | Vector output directory |
| `PASTA_RAS` | `PASTA_ZIPS\raster` | Raster output directory |
| `CRS_ALVO` | `ESRI:102033` | Target CRS |
| `TILE_M` | `25000` | Tile size in metres |
| `MAX_WORKERS` | `4` | Parallel threads |

**Usage:**
```bash
python 03_convert_tiled_vectors_rasters.py
```

---

### `04_deforestation_chart.py`

Generates a publication-ready bar chart of annual Amazon deforestation rates (INPE/PRODES), including project targets.

**Output:** `amazon_deforestation_norad.png`

**Usage:**
```bash
python 04_deforestation_chart.py
```

---

## Dependencies

All scripts auto-install their dependencies via `pip` on first run. For a manual install:

```bash
# Core (scripts 01–03)
pip install requests beautifulsoup4 lxml tqdm geopandas pyogrio pyarrow shapely rasterio numpy

# Script 04 only
pip install matplotlib numpy

# Optional — Selenium fallback for script 01
pip install selenium webdriver-manager
```

Python **3.11+** recommended.

---

## Directory Layout

```
C:\Amintas\Prodes\
├── zip\
│   └── YYYY-MM-DD\          ← date-stamped download folders
│       ├── <Biome>\
│       │   └── <Category>\
│       │       └── *.zip
│       ├── geoparquet\      ← vector tiles (script 03)
│       ├── raster\          ← COG GeoTIFF tiles (script 03)
│       ├── terrabrasilis_zips.csv
│       ├── terrabrasilis_zips.json
│       └── validation_report.json
└── geoparquet\              ← DuckDB-optimized output (script 02)
    └── report_YYYYMMDD_HHMMSS.json
```

---

## License

MIT
