# PRODES Data Pipeline

A collection of Python scripts to download, convert, visualize, and present
deforestation data from INPE's [TerraBrasilis](https://terrabrasilis.dpi.inpe.br/en/download-files/) platform (PRODES).

**Author:** Amintas Brandão Jr. \<abrandaojr@gmail.com\>  
**Affiliation:** Imazon — Instituto do Homem e Meio Ambiente da Amazônia

---

## Pipeline Overview

```
TerraBrasilis website
        │
        ▼
01_download_zips.py          ← scrape + download all .zip files
        │                       fallback: opens browser if scraping fails
        ▼  (ZIP archives)
02_convert_to_geoparquet.py  ← convert all formats in one pass
        │  Vectors (SHP/GPKG) → GeoParquet (Hilbert sort, zstd)
        │  Rasters (TIF/TIFF) → COG GeoTIFF (ESRI:102033, zonal-stats optimized)
        │
        ▼  (analysis-ready files)
03_deforestation_chart.py    ← annual deforestation rate chart (PNG)
        │
        ▼
04_generate_presentation.py  ← bilingual press PowerPoint (PT-BR + EN-US)
```

Run the full pipeline with a single command:

```bash
python 00_pipeline.py              # all steps, stop on first failure
python 00_pipeline.py -k           # all steps, continue past failures
python 00_pipeline.py --from 2     # resume from step 2
python 00_pipeline.py --steps 1 3  # run only steps 1 and 3
python 00_pipeline.py --from 2 -k  # resume from step 2, keep going on error
```

---

## Scripts

### `01_download_zips.py`

Discovers, downloads, and validates all `.zip` files from TerraBrasilis.

**Features:**
- Static HTML scraping with `lxml`/`html.parser` fallback
- Selenium fallback if JS rendering is required (Chrome must be installed)
- If both scraping methods fail, **opens the download page in your default browser**
- Inventory table: already-downloaded vs. pending files
- Parallel-safe resume support — interrupted downloads continue from where they stopped
- ZIP integrity check + automatic repair loop (up to 3 attempts per file)
- Permanent skip list via `CONFIG["skip_files"]`
- Saves `validation_report.json` and a CSV/JSON index of all files found

**Configuration** (`CONFIG` dict at the top of the file):

| Key | Default | Description |
|---|---|---|
| `base_url` | TerraBrasilis download page | Source URL to scrape |
| `root_folder` | `C:\Amintas\Prodes\zip` | Root storage directory |
| `download_timeout` | `600` s | Per-file download timeout |
| `chunk_size` | `32 MB` | Stream chunk size (maximizes throughput) |
| `skip_files` | `["prodes_brasil_2023_arte.zip"]` | Files to permanently skip |

```bash
python 01_download_zips.py
```

---

### `02_convert_to_geoparquet.py`

Converts all PRODES ZIP archives in a single pass. Handles both vectors and rasters.

**Vectors (SHP / GPKG)** → GeoParquet
- Hilbert curve row ordering via [`vector-to-geoparquet`](https://github.com/abrandaojr/vector-to-geoparquet)
- zstd compression, configurable row group size
- Incremental: skips already-converted files

**Rasters (TIF / TIFF)** → Cloud-Optimized GeoTIFF
- Reprojects to **ESRI:102033** (South America Equidistant Conic — equal-area, required for correct zonal stats)
- 512×512 internal tiles and DEFLATE+predictor=2 (optimal for CPU-bound zonal stats on 64 GB RAM)
- Internal overviews ×2, ×4, ×8, ×16, ×32

**Configuration** (`CONFIG` dict at the top of the file):

| Key | Default | Description |
|---|---|---|
| `source_dir` | `None` (auto-detects most recent dated folder) | Directory with input ZIPs |
| `dest_dir` | `C:\Amintas\Prodes\geoparquet` | Output directory |
| `n_workers` | `8` | Parallel conversion threads |
| `tile_size_m` | `25 000` | Hilbert sort tile granularity (m) |
| `compression` | `zstd` | Vector Parquet compression |
| `hilbert_p` | `15` | Hilbert curve precision (2^p grid) |
| `raster_crs` | `ESRI:102033` | Target CRS for all rasters |
| `cog_tile_px` | `512` | Internal COG tile size (px) |
| `raster_compress` | `DEFLATE` | Raster compression codec |
| `overview_levels` | `[2,4,8,16,32]` | Overview decimation factors |

```bash
python 02_convert_to_geoparquet.py           # convert
python 02_convert_to_geoparquet.py --list    # list existing outputs with CRS info
```

---

### `03_deforestation_chart.py`

Generates a publication-ready bar chart of annual Amazon deforestation rates
(INPE/PRODES), including historical data and 2028 project targets.

**Output:** `amazon_deforestation_norad.png`

**Configuration** (`CONFIG` dict at the top of the file):

| Key | Default | Description |
|---|---|---|
| `output_path` | `amazon_deforestation_norad.png` | Output file path |
| `dpi` | `220` | Image resolution |

```bash
python 03_deforestation_chart.py
```

---

### `04_generate_presentation.py`

Generates a 20-slide bilingual PowerPoint for press briefings on PRODES data.
Slides 1–10 in PT-BR; slides 11–20 in EN-US.

**Slides per section:**

| # | Content | Visual |
|---|---|---|
| 1 / 11 | Cover | Title + color bar |
| 2 / 12 | Lead stat: 56% decline | Large number card |
| 3 / 13 | Amazon historical series 2015–2028 | Bar chart (NYT style) |
| 4 / 14 | Deforestation by biome (2023) | Horizontal bars |
| 5 / 15 | Cerrado spotlight | Two stat cards |
| 6 / 16 | Forest cover remaining % | Horizontal bars (color-coded) |
| 7 / 17 | 2028 target trajectory | Bars + dashed projection line |
| 8 / 18 | International comparison (GFW/FAO) | Horizontal bars |
| 9 / 19 | Drivers & risks | Text with ▲▼ icons |
| 10 / 20 | Key takeaways | 3 numbered cards |

**Output:** `PRODES_Press_Briefing.pptx`

**Data sources:** All PRODES statistics calculated on-the-fly from GeoParquet files generated by script 02. Reference data (forest cover % from MapBiomas, international comparison from GFW/FAO) kept as labeled constants.

**Configuration** (`CONFIG` dict at the top of the file):

| Key | Default | Description |
|---|---|---|
| `output_path` | `PRODES_Press_Briefing.pptx` | Output file |
| `chart_dpi` | `220` | Chart image resolution |

```bash
python 04_generate_presentation.py
```

---

## Dependencies

All scripts **auto-install their dependencies** on first run via `pip`/`uv`.
For a manual install:

```bash
# Recommended: use uv for faster installs and better Windows wheel resolution
pip install uv
uv pip install -r requirements.txt
```

Python **3.11+** recommended.

---

## Directory Layout

```
C:\Amintas\Prodes\
├── zip\
│   └── YYYY-MM-DD\                   ← date-stamped download folders (script 01)
│       ├── <Biome>\
│       │   └── <Category>\
│       │       └── *.zip
│       ├── terrabrasilis_zips.csv
│       ├── terrabrasilis_zips.json
│       └── validation_report.json
└── geoparquet\                        ← converted outputs (script 02)
    ├── <rel_dir>\
    │   └── <zip_stem>\
    │       ├── <gpkg_stem>\
    │       │   └── <layer>.parquet    ← vector layers (Hilbert sorted)
    │       ├── <shp_stem>.parquet     ← orphan shapefiles
    │       └── <tif_stem>.tif         ← COG GeoTIFF (ESRI:102033)
    └── report_YYYYMMDD_HHMMSS.json

scripts\
├── amazon_deforestation_norad.png     ← output of script 03
└── PRODES_Press_Briefing.pptx         ← output of script 04
```

---

## License

MIT
