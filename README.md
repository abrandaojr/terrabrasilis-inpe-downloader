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
01_download_zips.py          ← scrape + download all .zip files
        │
        ▼  (ZIP archives)
02_convert_to_geoparquet.py  ← convert all formats in one pass
        │  Vectors (SHP/GPKG) → GeoParquet (Hilbert sort, zstd)
        │  Rasters (TIF/TIFF) → COG GeoTIFF (ESRI:102033, optimized for zonal stats)
        │
        ▼  (analysis-ready files)
04_deforestation_chart.py    ← generate annual deforestation rate chart
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
- Saves `validation_report.json` and a CSV/JSON index of all files found

**Configuration** (constants at the top of the file):

| Variable | Default | Description |
|---|---|---|
| `BASE_URL` | TerraBrasilis download page | Source URL to scrape |
| `ROOT_FOLDER` | `C:\Amintas\Prodes\zip` | Root storage directory |
| `DEST_FOLDER` | `ROOT_FOLDER\<today>` | Download destination |

```bash
python 01_download_zips.py
```

---

### `02_convert_to_geoparquet.py`

Converts all PRODES ZIP archives in a single pass. Handles both vectors and rasters.

**Vectors (SHP / GPKG)** → GeoParquet
- DuckDB-optimized: Hilbert curve row ordering, zstd compression
- Supports all GPKG layers; orphan SHPs as fallback
- Incremental: skips already-converted files

**Rasters (TIF / TIFF)** → Cloud-Optimized GeoTIFF
- Reprojects to **ESRI:102033** (South America Equidistant Conic — equal-area, required for correct zonal stats area calculations)
- COG format with 512×512 internal tiles and DEFLATE+predictor=2 compression
- Internal overviews at ×2, ×4, ×8, ×16, ×32 — lets `rasterstats` and similar tools pick the right resolution automatically without loading the full file
- Optimized for machines with 64 GB RAM: tile size chosen to balance sequential throughput and random-access reads from small polygons

**Configuration** (the `CONFIG` dict at the top of the file):

| Key | Default | Description |
|---|---|---|
| `source_dir` | `C:\Amintas\Prodes\zip\2026-05-07` | Directory with input ZIPs |
| `dest_dir` | `C:\Amintas\Prodes\geoparquet` | Output directory |
| `n_workers` | `8` | Parallel conversion threads |
| `compression` | `zstd` | Vector Parquet compression |
| `hilbert_p` | `15` | Hilbert sort precision |
| `raster_crs` | `ESRI:102033` | Target CRS for rasters |
| `cog_tile_px` | `512` | Internal COG tile size (px) |
| `raster_compress` | `DEFLATE` | Raster compression codec |
| `overview_levels` | `[2,4,8,16,32]` | Overview decimation factors |

```bash
python 02_convert_to_geoparquet.py           # convert
python 02_convert_to_geoparquet.py --list    # list existing outputs with CRS info
```

---

### `04_deforestation_chart.py`

Generates a publication-ready bar chart of annual Amazon deforestation rates (INPE/PRODES), including project targets.

**Output:** `amazon_deforestation_norad.png`

```bash
python 04_deforestation_chart.py
```

---

## Dependencies

All scripts auto-install their dependencies via `pip` on first run. For a manual install:

```bash
# Scripts 01 and 02
pip install requests beautifulsoup4 lxml tqdm geopandas pyogrio pyarrow shapely rasterio numpy

# Script 04
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
│   └── YYYY-MM-DD\          ← date-stamped download folders (script 01)
│       ├── <Biome>\
│       │   └── <Category>\
│       │       └── *.zip
│       ├── terrabrasilis_zips.csv
│       ├── terrabrasilis_zips.json
│       └── validation_report.json
└── geoparquet\              ← converted outputs (script 02)
    ├── <rel_dir>\
    │   └── <zip_stem>\
    │       ├── <gpkg_stem>\
    │       │   └── <layer>.parquet   ← vector layers (Hilbert sorted)
    │       ├── <shp_stem>.parquet    ← orphan shapefiles
    │       └── <tif_stem>.tif        ← COG GeoTIFF (ESRI:102033)
    └── report_YYYYMMDD_HHMMSS.json
```

---

## License

MIT
