from __future__ import annotations

__version__ = "2.0.0"
__all__: list[str] = []

import importlib.util
import json
import logging
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from collections import defaultdict
from concurrent.futures import Future, ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
from typing import NamedTuple

logging.basicConfig(format="%(levelname)-8s %(message)s", level=logging.WARNING)
log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Dependency bootstrap
# ---------------------------------------------------------------------------


def _bootstrap(*packages: tuple[str, str]) -> None:
    """Install missing packages into the current Python environment.

    Strategy order (most to least reliable for targeting sys.executable):
      1. python -m pip          â€” always installs into the running interpreter
      2. uv pip --python        â€” faster wheel resolution for native libs
      3. python -m uv pip       â€” uv via module, same target guarantee
      4. pip --break-system-pkg â€” last resort for externally-managed envs

    After each attempt, importlib.invalidate_caches() re-scans site-packages
    so that newly installed packages are immediately discoverable.
    Only packages that remain missing are retried with subsequent strategies.
    """
    import importlib
    import shutil

    mod_by_pip = {pip: mod for pip, mod in packages}

    def _still_missing(pkgs: list[str]) -> list[str]:
        importlib.invalidate_caches()
        return [p for p in pkgs if not importlib.util.find_spec(mod_by_pip[p])]

    missing = _still_missing(list(mod_by_pip))
    if not missing:
        return

    if not shutil.which("uv"):
        subprocess.call(
            [sys.executable, "-m", "pip", "install", "--quiet", "uv"],
            stderr=subprocess.DEVNULL,
        )

    strategies = [
        [sys.executable, "-m", "pip", "install", "--quiet"],
        ["uv", "pip", "install", "--python", sys.executable, "--quiet"],
        [sys.executable, "-m", "uv", "pip", "install", "--python", sys.executable, "--quiet"],
        [sys.executable, "-m", "pip", "install", "--quiet", "--break-system-packages"],
    ]
    for base in strategies:
        if not missing:
            return
        try:
            subprocess.check_call(base + missing, stderr=subprocess.DEVNULL)
            missing = _still_missing(missing)
        except (subprocess.CalledProcessError, FileNotFoundError):
            pass

    if missing:
        sys.exit(f"[FATAL] Could not install: {' '.join(missing)}")


_bootstrap(
    ("git+https://github.com/abrandaojr/vector-to-geoparquet.git", "vector_to_geoparquet"),
    ("geopandas", "geopandas"),
    ("pyogrio", "pyogrio"),
    ("pyarrow", "pyarrow"),
    ("shapely", "shapely"),
    ("numpy", "numpy"),
    ("tqdm", "tqdm"),
    ("rasterio", "rasterio"),
)

# Import heavy packages at module level so the one-time cold-start
# compilation happens here (expected) rather than mid-scan inside a loop.
import numpy as np  # noqa: E402
import geopandas as _gpd  # noqa: E402 (unused directly, but triggers geopandas init)
import pyogrio  # noqa: E402 (used by _gpkg_layers; pre-load avoids lazy-import hang)
from tqdm import tqdm  # noqa: E402

from prodes_pipeline.data_quality import (
    LineageRecord,
    StageTimer,
    atomic_write_json,
    configure_json_logging,
    file_inventory,
    freshness_metrics,
    parquet_quality_profile,
    require_existing_dir,
    to_jsonable,
    write_run_report,
)
from prodes_pipeline.pipeline_contracts import GEOPARQUET_CONTRACT, ZIP_ARCHIVE_CONTRACT
from prodes_pipeline.config import GEOPARQUET_DIR, REPORTS_DIR, ZIP_ROOT, ensure_pipeline_dirs


# ---------------------------------------------------------------------------
# CONFIG  â† the only section that needs to be edited
# ---------------------------------------------------------------------------

CONFIG: dict[str, object] = {
    # ---- I/O ---------------------------------------------------------------
    # Root directory that contains all dated download sub-folders.
    # All *.zip files found recursively are used; when the same filename
    # appears in multiple dated folders, the most recently modified copy wins.
    "zip_root": ZIP_ROOT,
    "dest_dir": GEOPARQUET_DIR,
    # ---- Extraction --------------------------------------------------------
    # Set to a path to keep extracted files between runs (faster re-runs).
    # None = use a temporary directory that is deleted after each run.
    "extract_dir": None,  # e.g. set PRODES_EXTRACT_DIR to keep extracted files
    # ---- Parallelism -------------------------------------------------------
    "n_workers": 8,
    # ---- Vector GeoParquet options -----------------------------------------
    "tile_size_m": 25_000,
    "row_group_size": 65_536,
    "compression": "zstd",
    "compression_level": 3,
    "hilbert_p": 15,
    # ---- Raster COG options (optimized for zonal stats, 64 GB RAM) ---------
    # Target CRS for all rasters (equal-area â€” required for correct area math)
    "raster_crs": "ESRI:102033",
    # Internal COG tile size in pixels. 512Ã—512 is the sweet spot for
    # windowed reads in zonal stats: large enough for sequential I/O,
    # small enough for random-access queries over small polygons.
    "cog_tile_px": 512,
    # DEFLATE + predictor=2 gives the best decompress speed for float rasters
    # on CPU-bound zonal stats workloads. Switch to "ZSTD" if you need better
    # compression ratios and your GDAL supports it.
    "raster_compress": "DEFLATE",
    # Overview decimation levels. With 30 m resolution, level 32 covers
    # ~1 km â€” enough for national-scale visualisation without loading full res.
    "overview_levels": [2, 4, 8, 16, 32],
}


# ---------------------------------------------------------------------------
# Module-level constants derived from CONFIG
# ---------------------------------------------------------------------------


def _latest_zips(zip_root: Path) -> list[Path]:
    """
    Scan zip_root recursively and return the most up-to-date copy of each ZIP.

    When the same filename exists in multiple sub-folders (e.g. different
    dated download runs), only the most recently modified file is kept.
    The returned list is sorted for deterministic processing order.
    """
    best: dict[str, tuple[float, Path]] = {}  # filename -> (mtime, path)
    for zp in zip_root.rglob("*.zip"):
        try:
            st = zp.stat()
        except OSError:
            continue
        current = best.get(zp.name)
        if current is None or st.st_mtime > current[0]:
            best[zp.name] = (st.st_mtime, zp)
    return sorted(path for _, path in best.values())


_ZIP_ROOT = Path(str(CONFIG["zip_root"]))
_DEST_DIR = Path(str(CONFIG["dest_dir"]))
# _SOURCE_DIR kept as alias so build_manifest (which uses it for rel_dir) works
_SOURCE_DIR = _ZIP_ROOT
_SHP_SIDECAR = frozenset({".shp", ".dbf", ".shx", ".prj", ".cpg", ".qpj", ".sbn", ".sbx"})
_RASTER_EXT = frozenset({".tif", ".tiff"})
_CRS_FALLBACKS = {
    # Some GDAL/PROJ builds do not ship the ESRI authority database entry.
    # This is ESRI:102033, South America Albers Equal Area Conic.
    "ESRI:102033": (
        "+proj=aea +lat_0=-32 +lon_0=-60 +lat_1=-5 +lat_2=-42 "
        "+x_0=0 +y_0=0 +datum=WGS84 +units=m +no_defs"
    ),
}
SEP = "=" * 65
DIV = "-" * 65
REPORT_DIR = REPORTS_DIR
OBS_LOG = configure_json_logging(REPORT_DIR / "observability.jsonl")
_LAYER_CACHE_PATH = REPORT_DIR / "02_gpkg_layer_cache.json"
_LAYER_CACHE_DIRTY = False


def _load_layer_cache() -> dict[str, dict[str, object]]:
    try:
        with _LAYER_CACHE_PATH.open("r", encoding="utf-8") as fh:
            payload = json.load(fh)
    except (OSError, json.JSONDecodeError):
        return {}
    return payload if isinstance(payload, dict) else {}


_LAYER_CACHE: dict[str, dict[str, object]] = _load_layer_cache()


def _save_layer_cache() -> None:
    if _LAYER_CACHE_DIRTY:
        atomic_write_json(_LAYER_CACHE_PATH, _LAYER_CACHE)


def _zip_preflight(zip_files: list[Path]) -> dict[str, object]:
    """Validate the selected source ZIPs before conversion starts."""
    empty = [str(p) for p in zip_files if p.stat().st_size <= 0]
    bad = [str(p) for p in zip_files if p.stat().st_size > 0 and not zipfile.is_zipfile(p)]
    duplicate_names: dict[str, int] = {}
    for p in zip_files:
        duplicate_names[p.name] = duplicate_names.get(p.name, 0) + 1
    duplicates = sorted(name for name, count in duplicate_names.items() if count > 1)
    return {
        "contract": to_jsonable(ZIP_ARCHIVE_CONTRACT),
        "zip_count": len(zip_files),
        "inventory": file_inventory(zip_files),
        "freshness": freshness_metrics(zip_files, ZIP_ARCHIVE_CONTRACT.freshness),
        "empty_count": len(empty),
        "bad_zip_count": len(bad),
        "duplicate_selected_names": duplicates,
        "empty": empty,
        "bad_zip": bad,
    }


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------


class Job(NamedTuple):
    """One conversion unit (a single vector layer or a raster file)."""

    priority: int
    kind: str  # "gpkg" | "shp" | "tif"
    zip_path: Path
    rel_dir: str
    zip_stem: str
    internal: str  # member path inside the zip
    layer: str | None  # GPKG layer name; None for shp / tif
    out_path: Path
    local_path: Path | None


# ---------------------------------------------------------------------------
# Stage 1 â€” Inspect ZIPs â†’ build manifest
# ---------------------------------------------------------------------------


def _gpkg_layers(zip_path: Path, internal: str) -> list[str]:
    global _LAYER_CACHE_DIRTY

    try:
        st = zip_path.stat()
    except OSError:
        return []

    cache_key = f"{zip_path.resolve()}::{internal}"
    cached = _LAYER_CACHE.get(cache_key)
    if (
        isinstance(cached, dict)
        and cached.get("size") == st.st_size
        and cached.get("mtime_ns") == st.st_mtime_ns
        and isinstance(cached.get("layers"), list)
    ):
        return [str(layer) for layer in cached["layers"]]

    uri = (
        f"/vsizip/{str(zip_path.resolve()).replace(chr(92), '/')}/"
        f"{internal.replace(chr(92), '/')}"
    )
    try:
        layers = [str(row[0]) for row in pyogrio.list_layers(uri)]
        _LAYER_CACHE[cache_key] = {
            "size": st.st_size,
            "mtime_ns": st.st_mtime_ns,
            "layers": layers,
        }
        _LAYER_CACHE_DIRTY = True
        return layers
    except KeyboardInterrupt:
        raise
    except Exception:  # Changed from BaseException to Exception
        log.warning(f"Failed to list layers for {zip_path.name}/{internal}", exc_info=True)
        return []


def build_manifest(zip_files: list[Path]) -> list[Job]:
    """Inspect zip_files and return the complete conversion manifest.

    Fast path: if a ZIP's output directory already exists and contains
    non-empty converted files, the ZIP is skipped entirely.  This avoids
    the slow pyogrio.list_layers() call for already-processed archives.
    """
    gpkg_jobs: list[Job] = []
    shp_jobs: list[Job] = []
    tif_jobs: list[Job] = []
    skipped = 0

    bar = tqdm(zip_files, desc="  scanning", unit="zip", ncols=80, leave=True)
    for zip_path in bar:
        bar.set_postfix_str(zip_path.name[:40])
        rel_dir = str(zip_path.parent.relative_to(_SOURCE_DIR))
        zip_stem = zip_path.stem

        # â”€â”€ Fast path: skip ZIPs whose output directory has converted files â”€â”€
        out_dir = _DEST_DIR / rel_dir / zip_stem
        if out_dir.exists() and any(
            p.stat().st_size > 0
            for p in out_dir.rglob("*")
            if p.suffix in (".parquet", ".tif")
        ):
            skipped += 1
            continue

        try:
            with zipfile.ZipFile(zip_path) as archive:
                names = archive.namelist()
        except zipfile.BadZipFile:
            log.warning(f"corrupt archive: {zip_path.name}")
            continue

        gpkg_stems: set[str] = set()

        # GPKG layers
        for internal in (n for n in names if n.lower().endswith(".gpkg")):
            stem = Path(internal).stem
            layers = _gpkg_layers(zip_path, internal)
            if not layers:
                log.warning(f"no layers: {zip_path.name}/{internal}")
                continue
            gpkg_stems.add(stem.lower())
            for layer in layers:
                gpkg_jobs.append(
                    Job(
                        priority=0,
                        kind="gpkg",
                        zip_path=zip_path,
                        rel_dir=rel_dir,
                        zip_stem=zip_stem,
                        internal=internal,
                        layer=layer,
                        out_path=_DEST_DIR / rel_dir / zip_stem / stem / f"{layer}.parquet",
                        local_path=None,
                    )
                )

        # Orphan SHP (not already covered by a GPKG with the same stem)
        for internal in (n for n in names if n.lower().endswith(".shp")):
            stem = Path(internal).stem
            if stem.lower() not in gpkg_stems:
                shp_jobs.append(
                    Job(
                        priority=1,
                        kind="shp",
                        zip_path=zip_path,
                        rel_dir=rel_dir,
                        zip_stem=zip_stem,
                        internal=internal,
                        layer=None,
                        out_path=_DEST_DIR / rel_dir / zip_stem / f"{stem}.parquet",
                        local_path=None,
                    )
                )

        # Rasters
        for internal in (n for n in names if Path(n).suffix.lower() in _RASTER_EXT):
            stem = Path(internal).stem
            tif_jobs.append(
                Job(
                    priority=2,
                    kind="tif",
                    zip_path=zip_path,
                    rel_dir=rel_dir,
                    zip_stem=zip_stem,
                    internal=internal,
                    layer=None,
                    out_path=_DEST_DIR / rel_dir / zip_stem / f"{stem}.tif",
                    local_path=None,
                )
            )

    bar.close()
    _save_layer_cache()
    if skipped:
        log.info(f"[scan] {skipped} ZIP(s) skipped â€” output already exists")
    return gpkg_jobs + shp_jobs + tif_jobs


# ---------------------------------------------------------------------------
# Stage 2 â€” Extract pending source files
# ---------------------------------------------------------------------------


def extract_pending(jobs: list[Job], tmp_dir: Path) -> list[Job]:
    """Extract source files for jobs into tmp_dir, skipping already-present files."""
    unique: dict[tuple[Path, str], Path] = {}
    by_zip: dict[Path, list[tuple[str, Path]]] = defaultdict(list)
    for job in jobs:
        key = (job.zip_path, job.internal)
        if key not in unique:
            dest = tmp_dir / job.rel_dir / job.zip_stem / job.internal
            unique[key] = dest
            by_zip[job.zip_path].append((job.internal, dest))

    pending = {k: v for k, v in unique.items() if not v.exists()}
    skipped = len(unique) - len(pending)
    if skipped:
        log.info(f"[extract] {skipped} file(s) already extracted â€” skipping")

    raw_extract_dir = tmp_dir / "_raw"
    bar = tqdm(total=len(pending), desc="  extracting", unit="file", ncols=80)

    for zip_path, entries in by_zip.items():
        pending_entries = [(internal, dest) for internal, dest in entries if not dest.exists()]
        if not pending_entries:
            continue
        with zipfile.ZipFile(zip_path) as zf:
            members = zf.namelist()
            for internal, dest in pending_entries:
                bar.set_postfix_str(Path(internal).name[:40])
                dest.parent.mkdir(parents=True, exist_ok=True)

                extracted_path = Path(zf.extract(internal, raw_extract_dir))
                shutil.move(str(extracted_path), dest)

                if internal.lower().endswith(".shp"):
                    stem = Path(internal).stem
                    parent = str(Path(internal).parent)
                    for name in members:
                        p = Path(name)
                        if (
                            p.stem == stem
                            and p.suffix.lower() in _SHP_SIDECAR
                            and str(p.parent) == parent
                        ):
                            sidecar_dest = dest.parent / p.name
                            if not sidecar_dest.exists():
                                extracted_sidecar_path = Path(zf.extract(name, raw_extract_dir))
                                shutil.move(str(extracted_sidecar_path), sidecar_dest)
                bar.update(1)

    bar.close()
    shutil.rmtree(raw_extract_dir, ignore_errors=True)  # Clean up the raw extract dir
    return [job._replace(local_path=unique[(job.zip_path, job.internal)]) for job in jobs]


# ---------------------------------------------------------------------------
# Stage 3 â€” Convert
# ---------------------------------------------------------------------------


def _source_mb(job: Job) -> float:
    """Calculate the source file size in MiB."""
    assert job.local_path is not None, "local_path must not be None at this stage"
    try:
        if job.kind == "shp":
            return sum(
                p.stat().st_size
                for p in job.local_path.parent.iterdir()
                if p.stem == job.local_path.stem and p.suffix.lower() in _SHP_SIDECAR
            ) / 1_048_576
        return job.local_path.stat().st_size / 1_048_576
    except OSError:
        return 0.0


def _convert_vector(job: Job) -> tuple[str, float, str | None]:
    """Convert a vector layer to GeoParquet."""
    from vector_to_geoparquet import convert_to_geoparquet

    src_mb = _source_mb(job)
    job.out_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        convert_to_geoparquet(
            input_path=str(job.local_path),
            output_path=str(job.out_path),
            layer=job.layer,
            tile_size_m=float(CONFIG["tile_size_m"]),
            row_group_size=int(CONFIG["row_group_size"]),
            compression=str(CONFIG["compression"]),
            compression_level=int(CONFIG["compression_level"]),
            hilbert_p=int(CONFIG["hilbert_p"]),
        )
        if not job.out_path.exists() or job.out_path.stat().st_size == 0:
            raise RuntimeError("output missing or empty")
        return "ok", src_mb, None
    except Exception as exc:
        job.out_path.unlink(missing_ok=True)
        return "error", src_mb, str(exc)


def _raster_dst_crs():
    """Resolve the configured raster CRS, including local fallbacks."""
    from rasterio.crs import CRS

    crs_text = str(CONFIG["raster_crs"]).strip()
    try:
        return CRS.from_user_input(crs_text)
    except Exception:
        fallback = _CRS_FALLBACKS.get(crs_text.upper())
        if fallback is None:
            raise
        return CRS.from_user_input(fallback)


def _convert_raster(job: Job) -> tuple[str, float, str | None]:
    """
    Reproject to ESRI:102033 and write a COG GeoTIFF optimised for zonal stats.

    Design notes (64 GB RAM):
    - 512Ã—512 internal tiles balance sequential throughput and random access.
      Larger tiles (e.g. 1024) speed up sequential reads but hurt small-polygon
      zonal stats because GDAL must decompress a larger block per read.
    - DEFLATE + predictor=2 decompresses faster than ZSTD on most CPUs for
      continuous-value rasters (vegetation indices, biomass, etc.), making it
      the better choice when zonal stats is the primary workload.
    - Overviews up to Ã—32 let tools like rasterstats pick the right resolution
      automatically, avoiding full-res reads for coarse summary statistics.
    """
    import rasterio
    from rasterio.enums import Resampling
    from rasterio.warp import calculate_default_transform, reproject

    src_mb = _source_mb(job)
    job.out_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        dst_crs = _raster_dst_crs()
        tile_px = int(CONFIG["cog_tile_px"])
        compress = str(CONFIG["raster_compress"])
        overview_levels: list[int] = CONFIG["overview_levels"]  # type: ignore[assignment]

        with rasterio.open(job.local_path) as src:
            transform, width, height = calculate_default_transform(
                src.crs, dst_crs, src.width, src.height, *src.bounds
            )
            nodata = src.nodata
            if nodata is None:
                # Use NaN for float dtypes, 0 for integer dtypes if nodata is not defined
                nodata = (
                    float("nan")
                    if np.issubdtype(np.dtype(src.dtypes[0]), np.floating)
                    else 0
                )

            profile = src.profile.copy()
            profile.update(
                {
                    "crs": dst_crs,
                    "transform": transform,
                    "width": width,
                    "height": height,
                    "driver": "GTiff",
                    "compress": compress,
                    "predictor": 2,  # delta predictor â€” effective for all numeric types
                    "tiled": True,
                    "blockxsize": tile_px,
                    "blockysize": tile_px,
                    "bigtiff": "IF_SAFER",
                    "nodata": nodata,
                }
            )

            with rasterio.open(job.out_path, "w", **profile) as dst:
                for band_idx in range(1, src.count + 1):
                    reproject(
                        source=rasterio.band(src, band_idx),
                        destination=rasterio.band(dst, band_idx),
                        src_transform=src.transform,
                        src_crs=src.crs,
                        dst_transform=transform,
                        dst_crs=dst_crs,
                        resampling=Resampling.nearest,
                    )

        # Build internal overviews so zonal-stats tools can use reduced
        # resolution automatically without loading the full raster.
        with rasterio.open(job.out_path, "r+") as ds:
            ds.build_overviews(overview_levels, Resampling.nearest)
            ds.update_tags(ns="rio_overview", resampling="nearest")

        if not job.out_path.exists() or job.out_path.stat().st_size == 0:
            raise RuntimeError("output missing or empty")
        return "ok", src_mb, None

    except Exception as exc:
        job.out_path.unlink(missing_ok=True)
        return "error", src_mb, str(exc)


def _run_job(job: Job) -> tuple[str, float, str | None]:
    """Run the appropriate conversion function for a job."""
    try:
        if job.kind == "tif":
            return _convert_raster(job)
        return _convert_vector(job)
    except Exception as exc:
        try:
            src_mb = _source_mb(job)
        except Exception:
            src_mb = 0.0
        job.out_path.unlink(missing_ok=True)
        return "error", src_mb, str(exc)


# ---------------------------------------------------------------------------
# Reporting helpers
# ---------------------------------------------------------------------------


def _file_mb(path: Path) -> float | None:
    """Get file size in MiB, or None if file does not exist."""
    try:
        return path.stat().st_size / 1_048_576
    except OSError:
        return None


def _fmt_duration(seconds: float) -> str:
    """Format duration in seconds to HhMms or Ms."""
    s = int(seconds)
    if s < 60:
        return f"{s}s"
    m, s = divmod(s, 60)
    if m < 60:
        return f"{m}m{s:02d}s"
    h, m = divmod(m, 60)
    return f"{h}h{m:02d}m"


def _print_report(rows: list[dict], totals: dict) -> None:
    """Print a formatted summary report of conversion jobs."""
    rows = sorted(rows, key=lambda r: r["n"])
    wb = max(max((len(r["biome"]) for r in rows), default=5), 5)
    wl = max(max((len(r["label"]) for r in rows), default=20), 20)
    sep = "=" * (6 + 6 + wb + wl + 10 + 10 + 8 + 8)
    div = "-" * len(sep)

    print(f"\n{sep}")
    print(
        f"  {'#':>4}  {'Kind':<5}  {'Biome':<{wb}}  {'Layer / File':<{wl}}"
        f"  {'Src MB':>8}  {'Out MB':>8}  {'%':>6}  Status"
    )
    print(div)
    for r in rows:
        tag = {"ok": "ok", "skipped": "skipped", "error": "ERROR"}.get(r["status"], r["status"])
        src = f"{r['src_mb']:>8.1f}" if r["src_mb"] is not None else "       -"
        out = f"{r['out_mb']:>8.1f}" if r["out_mb"] is not None else "       -"
        pct = f"{r['pct']:>6.1f}" if r["pct"] is not None else "     -"
        print(
            f"  {r['n']:>4}  {r['kind']:<5}  {r['biome']:<{wb}}  {r['label']:<{wl}}"
            f"  {src}  {out}  {pct}  {tag}"
        )
    print(sep)

    ok_rows = [r for r in rows if r["status"] != "error"]
    tot_src = sum(r["src_mb"] or 0 for r in ok_rows)
    tot_out = sum(r["out_mb"] or 0 for r in ok_rows)
    tot_pct = (tot_out / tot_src * 100) if tot_src else 0.0
    print(
        f"  Total : {totals['total']}  "
        f"ok: {totals['ok']}  "
        f"skipped: {totals['skipped']}  "
        f"errors: {totals['errors']}  |  "
        f"Src: {tot_src:.1f} MiB  â†’  Out: {tot_out:.1f} MiB  ({tot_pct:.1f}%)  "
        f"elapsed: {_fmt_duration(totals['elapsed'])}"
    )
    print(sep)


# ---------------------------------------------------------------------------
# --list command
# ---------------------------------------------------------------------------


def list_outputs() -> None:
    """Print a table of every GeoParquet and COG GeoTIFF in dest_dir."""
    import pyarrow.parquet as pq
    import rasterio

    if not _DEST_DIR.exists():
        print(f"  Output directory not found: {_DEST_DIR}")
        return

    rows: list[dict] = []

    for p in sorted(_DEST_DIR.rglob("*.parquet")):
        try:
            meta = pq.ParquetFile(p).schema_arrow.metadata or {}
            geo = json.loads(meta[b"geo"]) if b"geo" in meta else {}
            crs_parts = []
            for col_meta in geo.get("columns", {}).values():
                crs = col_meta.get("crs")
                if isinstance(crs, dict):
                    aid = crs.get("id", {})
                    code = f"{aid.get('authority','')}:{aid.get('code','')}"
                    name = crs.get("name", "")
                    crs_parts.append(f"{code} ({name})" if code.strip(":") else name)
                elif crs is None:
                    crs_parts.append("null")
                else:
                    crs_parts.append(str(crs)[:50])
            info = "; ".join(crs_parts) or "no geo metadata"
            mb = p.stat().st_size / 1_048_576
        except Exception as exc:
            info, mb = f"error: {exc}", 0.0
        rows.append({"kind": "parquet", "path": str(p.relative_to(_DEST_DIR)), "mb": mb, "info": info})

    for p in sorted(_DEST_DIR.rglob("*.tif")):
        try:
            with rasterio.open(p) as ds:
                info = (
                    f"CRS: {ds.crs.to_string()[:40]}  "
                    f"bands: {ds.count}  res: {ds.transform.a:.1f} m  "
                    f"overviews: {ds.overviews(1)}"
                )
        except Exception as exc:
            info = f"error: {exc}"
        rows.append({"kind": "tif", "path": str(p.relative_to(_DEST_DIR)), "mb": p.stat().st_size / 1_048_576, "info": info})

    if not rows:
        print(f"  No outputs found in {_DEST_DIR}")
        return

    wp = min(max(len(r["path"]) for r in rows), 80)
    wi = max(max(len(r["info"]) for r in rows), 10)
    sep = "=" * (wp + wi + 26)
    print(f"\n{sep}")
    print(f"  {'Kind':<8}  {'File':<{wp}}  {'MB':>7}  Info")
    print("-" * len(sep))
    for r in rows:
        print(f"  {r['kind']:<8}  {r['path']:<{wp}}  {r['mb']:>7.1f}  {r['info']}")
    print(sep)

    n_pq = sum(1 for r in rows if r["kind"] == "parquet")
    n_tif = sum(1 for r in rows if r["kind"] == "tif")
    total = sum(r["mb"] for r in rows)
    print(f"  Files: {len(rows)}  ({n_pq} parquet, {n_tif} tif)  |  Total: {total:.1f} MB")
    print(sep)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main() -> None:
    """Main function to orchestrate the conversion process."""
    ensure_pipeline_dirs()
    t0 = time.perf_counter()
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"\n{SEP}")
    print(f"  PRODES â†’ GeoParquet + COG  v{__version__}  |  {now}")
    print(f"{SEP}")

    require_existing_dir(_ZIP_ROOT, "ZIP root")

    # 1. Existing outputs â€” indexed by filename so files moved to _organized/ by
    # 05_organize_geoparquet.py are still recognised (different path, same name).
    existing_names: set[str] = set()
    if _DEST_DIR.exists():
        existing_names = {
            p.name
            for pat in ("*.parquet", "*.tif")
            for p in _DEST_DIR.rglob(pat)
            if p.stat().st_size > 0
        }
    log.info(f"\n  [1/3] existing outputs : {len(existing_names)}")

    # 2. Discover ZIPs â€” all sub-folders of zip_root, most recent copy per filename
    zip_files = _latest_zips(_ZIP_ROOT)
    log.info(f"  [2/3] scanning {len(zip_files)} zip archive(s) from {_ZIP_ROOT} ...")
    if not zip_files:
        sys.exit("  No zip archives found â€” nothing to do.")

    preflight_timer = StageTimer("02_zip_preflight")
    zip_quality = _zip_preflight(zip_files)
    OBS_LOG.emit(
        "stage_metrics",
        **to_jsonable(
            preflight_timer.finish(
                "ok"
                if not zip_quality["empty_count"] and not zip_quality["bad_zip_count"]
                else "failed",
                input_row_count=len(zip_files),
                output_row_count=zip_quality["zip_count"],
                anomalies={
                    "freshness": zip_quality["freshness"].get("stale", []),
                    "schema": zip_quality["bad_zip"],
                    "volume": zip_quality["empty"],
                },
            )
        ),
    )
    if zip_quality["empty_count"] or zip_quality["bad_zip_count"]:
        report_path = write_run_report(
            REPORT_DIR,
            Path(__file__).name,
            {
                "status": "failed",
                "reason": "invalid source ZIPs",
                "zip_root": str(_ZIP_ROOT),
                "zip_quality": zip_quality,
            },
        )
        print(f"  Quality report: {report_path}")
        sys.exit(
            "[FATAL] Source ZIP preflight failed: "
            f"{zip_quality['empty_count']} empty, {zip_quality['bad_zip_count']} invalid."
        )

    manifest_timer = StageTimer("02_build_conversion_manifest")
    manifest = build_manifest(zip_files)
    n_gpkg = sum(1 for j in manifest if j.kind == "gpkg")
    n_shp = sum(1 for j in manifest if j.kind == "shp")
    n_tif = sum(1 for j in manifest if j.kind == "tif")
    log.info(
        f"        manifest: {len(manifest)} job(s)  "
        f"(gpkg layers: {n_gpkg}  orphan shapefiles: {n_shp}  rasters: {n_tif})"
    )
    OBS_LOG.emit(
        "stage_metrics",
        **to_jsonable(
            manifest_timer.finish(
                "ok",
                input_row_count=len(zip_files),
                output_row_count=len(manifest),
            )
        ),
    )

    # 3. Cross-reference by filename â€” catches files reorganised to _organized/ by
    # 05_organize_geoparquet.py that no longer live at the expected out_path.
    done_jobs = [j for j in manifest if j.out_path.name in existing_names]
    todo_jobs = [j for j in manifest if j.out_path.name not in existing_names]
    log.info(f"  [3/3] pending: {len(todo_jobs)}  |  already converted: {len(done_jobs)}")

    if not todo_jobs:
        print("\n  All files already converted.\n")
        output_files = (
            [p for pat in ("*.parquet", "*.tif") for p in _DEST_DIR.rglob(pat)]
            if _DEST_DIR.exists()
            else []
        )
        report_path = write_run_report(
            REPORT_DIR,
            Path(__file__).name,
            {
                "status": "ok",
                "action": "already_converted",
                "zip_root": str(_ZIP_ROOT),
                "dest_dir": str(_DEST_DIR),
                "manifest_count": len(manifest),
                "zip_quality": zip_quality,
                "output_quality": {
                    "contract": to_jsonable(GEOPARQUET_CONTRACT),
                    "inventory": file_inventory(output_files),
                    "freshness": freshness_metrics(
                        output_files, GEOPARQUET_CONTRACT.freshness
                    ),
                    "parquet_profile": parquet_quality_profile(
                        [p for p in output_files if p.suffix.lower() == ".parquet"],
                        GEOPARQUET_CONTRACT,
                    ),
                },
                "lineage": LineageRecord(
                    stage_name="02_convert_to_geoparquet",
                    upstream_sources=[str(p) for p in zip_files],
                    transformation="Inspect ZIP archives, extract vector/raster members, convert vectors to GeoParquet and rasters to COG GeoTIFF.",
                    downstream_outputs=[str(_DEST_DIR)],
                    contracts=[ZIP_ARCHIVE_CONTRACT.name, GEOPARQUET_CONTRACT.name],
                ),
            },
        )
        print(f"  Quality report: {report_path}")
        return

    # 4 + 5. Extract â†’ Convert
    _extract_cfg = CONFIG.get("extract_dir")
    if _extract_cfg:
        extract_root = Path(str(_extract_cfg))
        extract_root.mkdir(parents=True, exist_ok=True)
        persistent = True
    else:
        extract_root = Path(tempfile.mkdtemp(prefix="prodes_"))
        persistent = False

    log.info(f"\n  extract : {extract_root}  ({'persistent' if persistent else 'temp'})")
    log.info(f"  output  : {_DEST_DIR}")
    log.info(f"  workers : {CONFIG['n_workers']}\n")

    try:
        todo_jobs = extract_pending(todo_jobs, extract_root)

        rows: list[dict] = [
            {
                "n": i,
                "kind": j.kind,
                "biome": j.rel_dir.replace("\\", "/").split("/")[0],
                "label": j.layer or Path(j.internal).stem,
                "src_mb": _source_mb(j) if j.local_path else None, # Recalculate if local path exists for done jobs
                "out_mb": _file_mb(j.out_path),
                "pct": None,
                "status": "skipped",
            }
            for i, j in enumerate(done_jobs, 1)
        ]
        counts: dict[str, int] = {"ok": 0, "skipped": len(done_jobs), "errors": 0}
        offset = len(done_jobs)

        bar = tqdm(
            total=len(todo_jobs),
            desc="  converting",
            unit="file",
            ncols=88,
            bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}]",
        )

        with ThreadPoolExecutor(max_workers=int(CONFIG["n_workers"])) as pool:
            futures: dict[Future, tuple[int, Job]] = {
                pool.submit(_run_job, job): (offset + i, job)
                for i, job in enumerate(todo_jobs, 1)
            }
            for fut in as_completed(futures):
                i, job = futures[fut]
                status, src_mb, err = fut.result()
                biome = job.rel_dir.replace("\\", "/").split("/")[0]
                label = job.layer or Path(job.internal).stem
                out_mb = _file_mb(job.out_path) if status != "error" else None
                pct = (out_mb / src_mb * 100) if (out_mb and src_mb) else None
                rows.append(
                    {
                        "n": i,
                        "kind": job.kind,
                        "biome": biome,
                        "label": label,
                        "src_mb": src_mb,
                        "out_mb": out_mb,
                        "pct": pct,
                        "status": status,
                    }
                )
                counts["ok" if status == "ok" else "errors" if status == "error" else "skipped"] += 1
                if status == "error":
                    log.error(f"  [ERROR] {label}: {err}")
                bar.update(1)

        bar.close()

    finally:
        if not persistent:
            shutil.rmtree(extract_root, ignore_errors=True)
            log.info(f"\n  temp removed: {extract_root}")
        else:
            log.info(f"\n  extracted files kept in: {extract_root}")

    elapsed = time.perf_counter() - t0
    _print_report(rows, {"total": len(manifest), **counts, "elapsed": elapsed})

    _DEST_DIR.mkdir(parents=True, exist_ok=True)
    output_files = [p for pat in ("*.parquet", "*.tif") for p in _DEST_DIR.rglob(pat)]
    output_quality = {
        "contract": to_jsonable(GEOPARQUET_CONTRACT),
        "inventory": file_inventory(output_files),
        "freshness": freshness_metrics(output_files, GEOPARQUET_CONTRACT.freshness),
        "parquet_profile": parquet_quality_profile(
            [p for p in output_files if p.suffix.lower() == ".parquet"],
            GEOPARQUET_CONTRACT,
        ),
    }
    OBS_LOG.emit(
        "data_contract",
        stage_name="02_geoparquet_output_contract",
        contract=to_jsonable(GEOPARQUET_CONTRACT),
        metrics=output_quality,
    )
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = _DEST_DIR / f"report_{ts}.json"
    atomic_write_json(
        report_path,
        {
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "version": __version__,
            "config": {k: str(v) for k, v in CONFIG.items()},
            "summary": {"total": len(manifest), **counts, "elapsed_s": round(elapsed, 1)},
            "zip_quality": zip_quality,
            "output_quality": output_quality,
            "lineage": LineageRecord(
                stage_name="02_convert_to_geoparquet",
                upstream_sources=[str(p) for p in zip_files],
                transformation="Convert TerraBrasilis ZIP archive members into analysis-ready GeoParquet and COG assets.",
                downstream_outputs=[str(_DEST_DIR)],
                contracts=[ZIP_ARCHIVE_CONTRACT.name, GEOPARQUET_CONTRACT.name],
            ),
            "jobs": [
                {
                    "n": r["n"],
                    "kind": r["kind"],
                    "biome": r["biome"],
                    "label": r["label"],
                    "status": r["status"],
                    "src_mb": round(r["src_mb"], 2) if r["src_mb"] is not None else None,
                    "out_mb": round(r["out_mb"], 2) if r["out_mb"] is not None else None,
                    "pct": round(r["pct"], 1) if r["pct"] is not None else None,
                }
                for r in sorted(rows, key=lambda r: r["n"])
            ],
        },
    )
    log.info(f"  report  : {report_path}\n")


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if cmd == "--list":
        list_outputs()
    else:
        main()

