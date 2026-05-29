"""
02_convert_to_geoparquet.py
===========================
Batch-convert PRODES ZIP archives to analysis-ready formats:

  Vectors (SHP / GPKG) → DuckDB-optimized GeoParquet (Hilbert sort, zstd)
  Rasters (TIF / TIFF) → Cloud-Optimized GeoTIFF reprojetado para ESRI:102033,
                          otimizado para zonal stats em máquinas com 64 GB RAM

Pipeline
--------
1. Scan source_dir for existing outputs (incremental — skips converted files).
2. Inspect every .zip: collect GPKG layers, orphan SHPs, and TIFs.
3. Cross-reference against existing outputs.
4. Extract pending source files to a managed temp directory.
5. Convert in parallel; clean up temp on exit or error.

Configuration
-------------
Edit the CONFIG dict below. All other symbols are implementation details.

Usage
-----
    python 02_convert_to_geoparquet.py           # convert
    python 02_convert_to_geoparquet.py --list    # list existing outputs

Author
------
Amintas Brandão Jr. <abrandaojr@gmail.com>
Imazon — Instituto do Homem e Meio Ambiente da Amazônia

License
-------
MIT
"""

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
from concurrent.futures import Future, ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path
from typing import NamedTuple

logging.basicConfig(format="%(levelname)-8s %(message)s", level=logging.WARNING)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Dependency bootstrap
# ---------------------------------------------------------------------------

def _bootstrap(*packages: tuple[str, str]) -> None:
    """Install missing packages at runtime.

    Tries uv first (better wheel resolution for native libs on Windows).
    Self-installs uv via pip if it is not found on PATH.
    Falls back to plain pip as a last resort.
    """
    import shutil

    missing = [pip for pip, mod in packages if not importlib.util.find_spec(mod)]
    if not missing:
        return

    if not shutil.which("uv"):
        subprocess.call(
            [sys.executable, "-m", "pip", "install", "--quiet", "uv"],
            stderr=subprocess.DEVNULL,
        )

    strategies = [
        ["uv", "pip", "install", "--quiet", *missing],
        [sys.executable, "-m", "uv", "pip", "install", "--quiet", *missing],
        [sys.executable, "-m", "pip", "install", "--quiet", *missing],
        [sys.executable, "-m", "pip", "install", "--quiet", "--break-system-packages", *missing],
    ]
    for cmd in strategies:
        try:
            subprocess.check_call(cmd, stderr=subprocess.DEVNULL)
            return
        except (subprocess.CalledProcessError, FileNotFoundError):
            pass
    sys.exit(f"[FATAL] Could not install: {' '.join(missing)}")


_bootstrap(
    ("git+https://github.com/abrandaojr/vector-to-geoparquet.git", "vector_to_geoparquet"),
    ("geopandas",            "geopandas"),
    ("pyogrio",              "pyogrio"),
    ("pyarrow",              "pyarrow"),
    ("shapely",              "shapely"),
    ("numpy",                "numpy"),
    ("tqdm",                 "tqdm"),
    ("rasterio",             "rasterio"),
)

from tqdm import tqdm  # noqa: E402

# ---------------------------------------------------------------------------
# CONFIG  ← the only section that needs to be edited
# ---------------------------------------------------------------------------

CONFIG: dict[str, object] = {
    # ---- I/O ---------------------------------------------------------------
    "source_dir": r"C:\Amintas\Prodes\zip\2026-05-07",
    "dest_dir":   r"C:\Amintas\Prodes\geoparquet",

    # ---- Extraction --------------------------------------------------------
    # Set to a path to keep extracted files between runs (faster re-runs).
    # None = use a temporary directory that is deleted after each run.
    "extract_dir": None,   # e.g. r"C:\Amintas\Prodes\extracted"

    # ---- Parallelism -------------------------------------------------------
    "n_workers": 8,

    # ---- Vector GeoParquet options -----------------------------------------
    "tile_size_m":       25_000,
    "row_group_size":    65_536,
    "compression":       "zstd",
    "compression_level": 3,
    "hilbert_p":         15,

    # ---- Raster COG options (optimized for zonal stats, 64 GB RAM) ---------
    # Target CRS for all rasters (equal-area — required for correct area math)
    "raster_crs":        "ESRI:102033",
    # Internal COG tile size in pixels. 512×512 is the sweet spot for
    # windowed reads in zonal stats: large enough for sequential I/O,
    # small enough for random-access queries over small polygons.
    "cog_tile_px":       512,
    # DEFLATE + predictor=2 gives the best decompress speed for float rasters
    # on CPU-bound zonal stats workloads. Switch to "ZSTD" if you need better
    # compression ratios and your GDAL supports it.
    "raster_compress":   "DEFLATE",
    # Overview decimation levels. With 30 m resolution, level 32 covers
    # ~1 km — enough for national-scale visualisation without loading full res.
    "overview_levels":   [2, 4, 8, 16, 32],
}

# ---------------------------------------------------------------------------
# Module-level constants derived from CONFIG
# ---------------------------------------------------------------------------

_SOURCE_DIR  = Path(str(CONFIG["source_dir"]))
_DEST_DIR    = Path(str(CONFIG["dest_dir"]))
_SHP_SIDECAR = frozenset({".shp", ".dbf", ".shx", ".prj", ".cpg", ".qpj", ".sbn", ".sbx"})
_RASTER_EXT  = frozenset({".tif", ".tiff"})
SEP          = "=" * 65
DIV          = "-" * 65

# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

class Job(NamedTuple):
    """One conversion unit (a single vector layer or a raster file)."""

    priority:   int
    kind:       str          # "gpkg" | "shp" | "tif"
    zip_path:   Path
    rel_dir:    str
    zip_stem:   str
    internal:   str          # member path inside the zip
    layer:      str | None   # GPKG layer name; None for shp / tif
    out_path:   Path
    local_path: Path | None

# ---------------------------------------------------------------------------
# Stage 1 — Inspect ZIPs → build manifest
# ---------------------------------------------------------------------------

def _gpkg_layers(zip_path: Path, internal: str) -> list[str]:
    import pyogrio
    uri = (
        f"/vsizip/{str(zip_path.resolve()).replace(chr(92), '/')}/"
        f"{internal.replace(chr(92), '/')}"
    )
    try:
        return [str(row[0]) for row in pyogrio.list_layers(uri)]
    except Exception:
        return []


def build_manifest(zip_files: list[Path]) -> list[Job]:
    """Inspect zip_files and return the complete conversion manifest."""
    gpkg_jobs: list[Job] = []
    shp_jobs:  list[Job] = []
    tif_jobs:  list[Job] = []

    bar = tqdm(zip_files, desc="  scanning", unit="zip", ncols=80, leave=True)
    for zip_path in bar:
        bar.set_postfix_str(zip_path.name[:40])
        rel_dir  = str(zip_path.parent.relative_to(_SOURCE_DIR))
        zip_stem = zip_path.stem

        try:
            names = zipfile.ZipFile(zip_path).namelist()
        except zipfile.BadZipFile:
            tqdm.write(f"  [WARN] corrupt archive: {zip_path.name}")
            continue

        gpkg_stems: set[str] = set()

        # GPKG layers
        for internal in (n for n in names if n.lower().endswith(".gpkg")):
            stem   = Path(internal).stem
            layers = _gpkg_layers(zip_path, internal)
            if not layers:
                tqdm.write(f"  [WARN] no layers: {zip_path.name}/{internal}")
                continue
            gpkg_stems.add(stem.lower())
            for layer in layers:
                gpkg_jobs.append(Job(
                    priority=0, kind="gpkg",
                    zip_path=zip_path, rel_dir=rel_dir, zip_stem=zip_stem,
                    internal=internal, layer=layer,
                    out_path=_DEST_DIR / rel_dir / zip_stem / stem / f"{layer}.parquet",
                    local_path=None,
                ))

        # Orphan SHP (not already covered by a GPKG with the same stem)
        for internal in (n for n in names if n.lower().endswith(".shp")):
            stem = Path(internal).stem
            if stem.lower() not in gpkg_stems:
                shp_jobs.append(Job(
                    priority=1, kind="shp",
                    zip_path=zip_path, rel_dir=rel_dir, zip_stem=zip_stem,
                    internal=internal, layer=None,
                    out_path=_DEST_DIR / rel_dir / zip_stem / f"{stem}.parquet",
                    local_path=None,
                ))

        # Rasters
        for internal in (n for n in names if Path(n).suffix.lower() in _RASTER_EXT):
            stem = Path(internal).stem
            tif_jobs.append(Job(
                priority=2, kind="tif",
                zip_path=zip_path, rel_dir=rel_dir, zip_stem=zip_stem,
                internal=internal, layer=None,
                out_path=_DEST_DIR / rel_dir / zip_stem / f"{stem}.tif",
                local_path=None,
            ))

    bar.close()
    return gpkg_jobs + shp_jobs + tif_jobs

# ---------------------------------------------------------------------------
# Stage 2 — Extract pending source files
# ---------------------------------------------------------------------------

def extract_pending(jobs: list[Job], tmp_dir: Path) -> list[Job]:
    """Extract source files for jobs into tmp_dir, skipping already-present files."""
    unique: dict[tuple[Path, str], Path] = {}
    for job in jobs:
        key = (job.zip_path, job.internal)
        if key not in unique:
            unique[key] = tmp_dir / job.rel_dir / job.zip_stem / job.internal

    pending = {k: v for k, v in unique.items() if not v.exists()}
    skipped = len(unique) - len(pending)
    if skipped:
        tqdm.write(f"  [extract] {skipped} file(s) already extracted — skipping")

    raw = tmp_dir / "_raw"
    bar = tqdm(total=len(pending), desc="  extracting", unit="file", ncols=80)

    for (zip_path, internal), dest in unique.items():
        if dest.exists():
            continue
        bar.set_postfix_str(Path(internal).name[:40])
        dest.parent.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(zip_path) as zf:
            members = zf.namelist()
            shutil.move(str(Path(zf.extract(internal, raw))), dest)
            if internal.lower().endswith(".shp"):
                stem   = Path(internal).stem
                parent = str(Path(internal).parent)
                for name in members:
                    p = Path(name)
                    if (p.stem == stem
                            and p.suffix.lower() in _SHP_SIDECAR
                            and str(p.parent) == parent):
                        sidecar_dest = dest.parent / p.name
                        if not sidecar_dest.exists():
                            shutil.move(str(Path(zf.extract(name, raw))), sidecar_dest)
        bar.update(1)

    bar.close()
    return [job._replace(local_path=unique[(job.zip_path, job.internal)]) for job in jobs]

# ---------------------------------------------------------------------------
# Stage 3 — Convert
# ---------------------------------------------------------------------------

def _source_mb(job: Job) -> float:
    try:
        if job.kind == "shp":
            return sum(
                p.stat().st_size
                for p in job.local_path.parent.iterdir()  # type: ignore[union-attr]
                if p.stem == job.local_path.stem           # type: ignore[union-attr]
                and p.suffix.lower() in _SHP_SIDECAR
            ) / 1_048_576
        return job.local_path.stat().st_size / 1_048_576  # type: ignore[union-attr]
    except OSError:
        return 0.0


def _convert_vector(job: Job) -> tuple[str, float, str | None]:
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


def _convert_raster(job: Job) -> tuple[str, float, str | None]:
    """
    Reproject to ESRI:102033 and write a COG GeoTIFF optimised for zonal stats.

    Design notes (64 GB RAM):
    - 512×512 internal tiles balance sequential throughput and random access.
      Larger tiles (e.g. 1024) speed up sequential reads but hurt small-polygon
      zonal stats because GDAL must decompress a larger block per read.
    - DEFLATE + predictor=2 decompresses faster than ZSTD on most CPUs for
      continuous-value rasters (vegetation indices, biomass, etc.), making it
      the better choice when zonal stats is the primary workload.
    - Overviews up to ×32 let tools like rasterstats pick the right resolution
      automatically, avoiding full-res reads for coarse summary statistics.
    """
    import numpy as np
    import rasterio
    from rasterio.crs import CRS
    from rasterio.enums import Resampling
    from rasterio.warp import calculate_default_transform, reproject

    src_mb   = _source_mb(job)
    dst_crs  = CRS.from_user_input(str(CONFIG["raster_crs"]))
    tile_px  = int(CONFIG["cog_tile_px"])
    compress = str(CONFIG["raster_compress"])
    overview_levels: list[int] = list(CONFIG["overview_levels"])  # type: ignore[arg-type]

    job.out_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        with rasterio.open(job.local_path) as src:
            transform, width, height = calculate_default_transform(
                src.crs, dst_crs, src.width, src.height, *src.bounds
            )
            nodata = src.nodata
            if nodata is None:
                nodata = float("nan") if np.issubdtype(np.dtype(src.dtypes[0]), np.floating) else 0

            profile = src.profile.copy()
            profile.update({
                "crs":        dst_crs,
                "transform":  transform,
                "width":      width,
                "height":     height,
                "driver":     "GTiff",
                "compress":   compress,
                "predictor":  2,          # delta predictor — effective for all numeric types
                "tiled":      True,
                "blockxsize": tile_px,
                "blockysize": tile_px,
                "bigtiff":    "IF_SAFER",
                "nodata":     nodata,
            })

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
    if job.kind == "tif":
        return _convert_raster(job)
    return _convert_vector(job)

# ---------------------------------------------------------------------------
# Reporting helpers
# ---------------------------------------------------------------------------

def _file_mb(path: Path) -> float | None:
    try:
        return path.stat().st_size / 1_048_576
    except OSError:
        return None


def _fmt_duration(seconds: float) -> str:
    s = int(seconds)
    if s < 60:
        return f"{s}s"
    m, s = divmod(s, 60)
    if m < 60:
        return f"{m}m{s:02d}s"
    h, m = divmod(m, 60)
    return f"{h}h{m:02d}m"


def _print_report(rows: list[dict], totals: dict) -> None:
    rows = sorted(rows, key=lambda r: r["n"])
    wb   = max(max((len(r["biome"]) for r in rows), default=5), 5)
    wl   = max(max((len(r["label"]) for r in rows), default=20), 20)
    sep  = "=" * (6 + 6 + wb + wl + 10 + 10 + 8 + 8)
    div  = "-" * len(sep)

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
        pct = f"{r['pct']:>6.1f}"    if r["pct"]    is not None else "     -"
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
        f"Src: {tot_src:.1f} MiB  →  Out: {tot_out:.1f} MiB  ({tot_pct:.1f}%)  "
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
            geo  = json.loads(meta[b"geo"]) if b"geo" in meta else {}
            crs_parts = []
            for col_meta in geo.get("columns", {}).values():
                crs = col_meta.get("crs")
                if isinstance(crs, dict):
                    aid  = crs.get("id", {})
                    code = f"{aid.get('authority','')}:{aid.get('code','')}"
                    name = crs.get("name", "")
                    crs_parts.append(f"{code} ({name})" if code.strip(":") else name)
                elif crs is None:
                    crs_parts.append("null")
                else:
                    crs_parts.append(str(crs)[:50])
            info = "; ".join(crs_parts) or "no geo metadata"
            mb   = p.stat().st_size / 1_048_576
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

    wp  = min(max(len(r["path"]) for r in rows), 80)
    wi  = max(max(len(r["info"]) for r in rows), 10)
    sep = "=" * (wp + wi + 26)
    print(f"\n{sep}")
    print(f"  {'Kind':<8}  {'File':<{wp}}  {'MB':>7}  Info")
    print("-" * len(sep))
    for r in rows:
        print(f"  {r['kind']:<8}  {r['path']:<{wp}}  {r['mb']:>7.1f}  {r['info']}")
    print(sep)

    n_pq  = sum(1 for r in rows if r["kind"] == "parquet")
    n_tif = sum(1 for r in rows if r["kind"] == "tif")
    total = sum(r["mb"] for r in rows)
    print(f"  Files: {len(rows)}  ({n_pq} parquet, {n_tif} tif)  |  Total: {total:.1f} MB")
    print(sep)

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    t0  = time.perf_counter()
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"\n{SEP}")
    print(f"  PRODES → GeoParquet + COG  v{__version__}  |  {now}")
    print(f"{SEP}")

    if not _SOURCE_DIR.exists():
        sys.exit(f"[FATAL] source directory not found: {_SOURCE_DIR}")

    # 1. Existing outputs
    existing: set[Path] = set()
    if _DEST_DIR.exists():
        existing = {
            p.resolve()
            for pat in ("*.parquet", "*.tif")
            for p in _DEST_DIR.rglob(pat)
            if p.stat().st_size > 0
        }
    print(f"\n  [1/3] existing outputs : {len(existing)}")

    # 2. Build manifest
    zip_files = sorted(_SOURCE_DIR.rglob("*.zip"))
    print(f"  [2/3] scanning {len(zip_files)} zip archive(s) ...")
    if not zip_files:
        sys.exit("  No zip archives found — nothing to do.")

    manifest = build_manifest(zip_files)
    n_gpkg   = sum(1 for j in manifest if j.kind == "gpkg")
    n_shp    = sum(1 for j in manifest if j.kind == "shp")
    n_tif    = sum(1 for j in manifest if j.kind == "tif")
    print(
        f"        manifest: {len(manifest)} job(s)  "
        f"(gpkg layers: {n_gpkg}  orphan shapefiles: {n_shp}  rasters: {n_tif})"
    )

    # 3. Cross-reference
    done_jobs = [j for j in manifest if j.out_path.resolve() in existing]
    todo_jobs = [j for j in manifest if j.out_path.resolve() not in existing]
    print(f"  [3/3] pending: {len(todo_jobs)}  |  already converted: {len(done_jobs)}")

    if not todo_jobs:
        print("\n  All files already converted.\n")
        return

    # 4 + 5. Extract → Convert
    _extract_cfg = CONFIG.get("extract_dir")
    if _extract_cfg:
        extract_root = Path(str(_extract_cfg))
        extract_root.mkdir(parents=True, exist_ok=True)
        persistent = True
    else:
        extract_root = Path(tempfile.mkdtemp(prefix="prodes_"))
        persistent = False

    print(f"\n  extract : {extract_root}  ({'persistent' if persistent else 'temp'})")
    print(f"  output  : {_DEST_DIR}")
    print(f"  workers : {CONFIG['n_workers']}\n")

    try:
        todo_jobs = extract_pending(todo_jobs, extract_root)

        rows: list[dict] = [
            {
                "n":      i,
                "kind":   j.kind,
                "biome":  j.rel_dir.replace("\\", "/").split("/")[0],
                "label":  j.layer or Path(j.internal).stem,
                "src_mb": None,
                "out_mb": _file_mb(j.out_path),
                "pct":    None,
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
                i, job              = futures[fut]
                status, src_mb, err = fut.result()
                biome  = job.rel_dir.replace("\\", "/").split("/")[0]
                label  = job.layer or Path(job.internal).stem
                out_mb = _file_mb(job.out_path) if status != "error" else None
                pct    = (out_mb / src_mb * 100) if (out_mb and src_mb) else None
                rows.append({
                    "n": i, "kind": job.kind, "biome": biome, "label": label,
                    "src_mb": src_mb, "out_mb": out_mb, "pct": pct, "status": status,
                })
                counts["ok" if status == "ok" else "errors" if status == "error" else "skipped"] += 1
                if status == "error":
                    tqdm.write(f"  [ERROR] {label}: {err}")
                bar.update(1)

        bar.close()

    finally:
        if not persistent:
            shutil.rmtree(extract_root, ignore_errors=True)
            print(f"\n  temp removed: {extract_root}")
        else:
            print(f"\n  extracted files kept in: {extract_root}")

    elapsed = time.perf_counter() - t0
    _print_report(rows, {"total": len(manifest), **counts, "elapsed": elapsed})

    _DEST_DIR.mkdir(parents=True, exist_ok=True)
    ts          = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = _DEST_DIR / f"report_{ts}.json"
    report_path.write_text(
        json.dumps(
            {
                "generated_at": datetime.now(timezone.utc).isoformat(),
                "version":      __version__,
                "config":       {k: str(v) for k, v in CONFIG.items()},
                "summary":      {"total": len(manifest), **counts, "elapsed_s": round(elapsed, 1)},
                "jobs": [
                    {
                        "n":      r["n"],
                        "kind":   r["kind"],
                        "biome":  r["biome"],
                        "label":  r["label"],
                        "status": r["status"],
                        "src_mb": round(r["src_mb"], 2) if r["src_mb"] is not None else None,
                        "out_mb": round(r["out_mb"], 2) if r["out_mb"] is not None else None,
                        "pct":    round(r["pct"],    1) if r["pct"]    is not None else None,
                    }
                    for r in sorted(rows, key=lambda r: r["n"])
                ],
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(f"  report  : {report_path}\n")


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if cmd == "--list":
        list_outputs()
    else:
        main()
