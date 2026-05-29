"""prodes_to_geoparquet.py
========================
Batch-convert PRODES zip archives to DuckDB-optimized GeoParquet files.

Pipeline
--------
1. Scan *dest_dir* for existing ``.parquet`` files (incremental runs).
2. Inspect every ``.zip`` in *source_dir* and build a full conversion manifest.
3. Cross-reference manifest against existing files; skip already-converted jobs.
4. Extract only the pending source files to a managed temporary directory.
5. Convert in parallel using a thread pool; clean up temp on exit or error.

Configuration
-------------
Edit the :data:`CONFIG` dict at the top of this file.  All other symbols are
considered private implementation details.

Usage
-----
.. code-block:: bash

    python prodes_to_geoparquet.py           # convert
    python prodes_to_geoparquet.py --list    # list existing GeoParquets

Author
------
Amintas Brandão Jr. <abrandaojr@gmail.com>
Imazon — Instituto do Homem e Meio Ambiente da Amazônia

License
-------
MIT
"""

from __future__ import annotations

__version__ = "1.1.0"
__all__: list[str] = []

import contextlib
import importlib
import importlib.util
import io
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
from typing import Callable, NamedTuple

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

logging.basicConfig(format="%(levelname)-8s %(message)s", level=logging.WARNING)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Dependency bootstrap
# ---------------------------------------------------------------------------

def _bootstrap(*packages: tuple[str, str]) -> None:
    """Install *packages* at runtime if any import spec is missing."""
    missing = [pip for pip, mod in packages if not importlib.util.find_spec(mod)]
    if not missing:
        return
    for extra in ([], ["--break-system-packages"]):
        cmd = [sys.executable, "-m", "pip", "install", "--quiet", *extra, *missing]
        try:
            subprocess.check_call(cmd, stderr=subprocess.DEVNULL)
            return
        except (subprocess.CalledProcessError, FileNotFoundError):
            pass
    sys.exit(f"[FATAL] Could not install: {' '.join(missing)}")


_bootstrap(
    ("requests",  "requests"),
    ("geopandas", "geopandas"),
    ("pyogrio",   "pyogrio"),
    ("pyarrow",   "pyarrow"),
    ("shapely",   "shapely"),
    ("numpy",     "numpy"),
    ("tqdm",      "tqdm"),
)

import requests  # noqa: E402
from tqdm import tqdm  # noqa: E402

# ---------------------------------------------------------------------------
# CONFIG  ← the only section that needs to be edited
# ---------------------------------------------------------------------------

CONFIG: dict[str, object] = {
    # ---- I/O ---------------------------------------------------------------
    "source_dir": r"C:\Amintas\Prodes\zip\2026-05-07",
    "dest_dir":   r"C:\Amintas\Prodes\geoparquet",

    # ---- Remote converter --------------------------------------------------
    "github_raw": (
        "https://raw.githubusercontent.com/abrandaojr/"
        "vector-to-geoparquet/main/vector_to_geoparquet.py"
    ),

    # ---- Extraction --------------------------------------------------------
    # Persistent directory for extracted source files.  If set, files are
    # kept between runs so re-runs skip already-extracted files.
    # If None, a temporary directory is created and deleted after each run.
    "extract_dir": None,   # e.g. r"C:\Amintas\Prodes\extracted"

    # ---- Parallelism -------------------------------------------------------
    "n_workers": 8,

    # ---- GeoParquet write options ------------------------------------------
    "tile_size_m":       25_000,
    "row_group_size":    65_536,
    "compression":       "zstd",
    "compression_level": 3,
    "hilbert_p":         15,
}

# ---------------------------------------------------------------------------
# Module-level constants derived from CONFIG (do not edit)
# ---------------------------------------------------------------------------

_SOURCE_DIR  = Path(str(CONFIG["source_dir"]))
_DEST_DIR    = Path(str(CONFIG["dest_dir"]))
_CACHE_MOD   = Path(__file__).parent / "_vector_to_geoparquet.py"
_GPQ_KWARGS: dict[str, object] = {
    k: CONFIG[k]
    for k in ("tile_size_m", "row_group_size", "compression",
              "compression_level", "hilbert_p")
}
_SHP_SIDECAR = frozenset({
    ".shp", ".dbf", ".shx", ".prj", ".cpg", ".qpj", ".sbn", ".sbx",
})

# ---------------------------------------------------------------------------
# Converter loader
# ---------------------------------------------------------------------------

def _load_converter() -> Callable:
    """Return ``convert_to_geoparquet``, fetching it from GitHub if needed."""
    if _CACHE_MOD.exists():
        text  = _CACHE_MOD.read_text(encoding="utf-8")
        stale = "__version__" not in text or '"1.2.0"' not in text
        if stale:
            _CACHE_MOD.unlink()
            print("  [cache] stale — re-fetching ...", end=" ", flush=True)
        else:
            print(f"  [cache] {_CACHE_MOD.name}")

    if not _CACHE_MOD.exists():
        print("  [fetch] downloading converter ...", end=" ", flush=True)
        response = requests.get(str(CONFIG["github_raw"]), timeout=30)
        response.raise_for_status()
        _CACHE_MOD.write_text(response.text, encoding="utf-8")
        print("ok")

    spec = importlib.util.spec_from_file_location("_vtgp", _CACHE_MOD)
    mod  = importlib.util.module_from_spec(spec)  # type: ignore[arg-type]
    mod.__name__ = "_vtgp"
    spec.loader.exec_module(mod)  # type: ignore[union-attr]
    return mod.convert_to_geoparquet  # type: ignore[attr-defined]

# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

class Job(NamedTuple):
    """One GeoParquet conversion unit."""

    priority:   int
    kind:       str
    zip_path:   Path
    rel_dir:    str
    zip_stem:   str
    internal:   str
    layer:      str | None
    out_path:   Path
    local_path: Path | None

# ---------------------------------------------------------------------------
# Stage 1 — Inspect
# ---------------------------------------------------------------------------

def _gpkg_layers(zip_path: Path, internal: str) -> list[str]:
    """Return layer names inside a GeoPackage member of a zip archive."""
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
    """Inspect *zip_files* and return the complete conversion manifest."""
    gpkg_jobs: list[Job] = []
    shp_jobs:  list[Job] = []

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

    bar.close()
    return gpkg_jobs + shp_jobs

# ---------------------------------------------------------------------------
# Stage 2 — Extract
# ---------------------------------------------------------------------------

def extract_pending(jobs: list[Job], tmp_dir: Path) -> list[Job]:
    """Extract source files for *jobs* into *tmp_dir*, skipping existing files."""
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
    """Return the size of the source file(s) in MiB."""
    try:
        if job.local_path.suffix.lower() == ".shp":  # type: ignore[union-attr]
            return sum(
                p.stat().st_size
                for p in job.local_path.parent.iterdir()  # type: ignore[union-attr]
                if p.stem == job.local_path.stem  # type: ignore[union-attr]
                and p.suffix.lower() in _SHP_SIDECAR
            ) / 1_048_576
        return job.local_path.stat().st_size / 1_048_576  # type: ignore[union-attr]
    except OSError:
        return 0.0


def _convert_job(job: Job, convert_fn: Callable) -> tuple[str, float, str | None]:
    """Execute one conversion job. Returns (status, src_mb, error)."""
    src_mb = _source_mb(job)
    job.out_path.parent.mkdir(parents=True, exist_ok=True)
    sink = io.StringIO()
    try:
        with contextlib.redirect_stdout(sink), contextlib.redirect_stderr(sink):
            convert_fn(
                input_path=str(job.local_path),
                output_path=str(job.out_path),
                layer=job.layer,
                **_GPQ_KWARGS,
            )
        if not job.out_path.exists() or job.out_path.stat().st_size == 0:
            raise RuntimeError("output file is missing or empty after conversion")
        return "ok", src_mb, None
    except Exception as exc:
        job.out_path.unlink(missing_ok=True)
        return "error", src_mb, str(exc)

# ---------------------------------------------------------------------------
# Reporting
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
        f"  {'Src MB':>8}  {'GPQ MB':>8}  {'%':>6}  Status"
    )
    print(div)
    for r in rows:
        tag = {"ok": "ok", "skipped": "skipped", "error": "ERROR"}.get(
            r["status"], r["status"]
        )
        src = f"{r['src_mb']:>8.1f}" if r["src_mb"] is not None else "       -"
        gpq = f"{r['gpq_mb']:>8.1f}" if r["gpq_mb"] is not None else "       -"
        pct = f"{r['pct']:>6.1f}"    if r["pct"]    is not None else "     -"
        print(
            f"  {r['n']:>4}  {r['kind']:<5}  {r['biome']:<{wb}}  {r['label']:<{wl}}"
            f"  {src}  {gpq}  {pct}  {tag}"
        )
    print(sep)

    ok_rows  = [r for r in rows if r["status"] != "error"]
    tot_src  = sum(r["src_mb"] or 0 for r in ok_rows)
    tot_gpq  = sum(r["gpq_mb"] or 0 for r in ok_rows)
    tot_pct  = (tot_gpq / tot_src * 100) if tot_src else 0.0
    print(
        f"  Total : {totals['total']}  "
        f"ok: {totals['ok']}  "
        f"skipped: {totals['skipped']}  "
        f"errors: {totals['errors']}  |  "
        f"Src: {tot_src:.1f} MiB  →  GPQ: {tot_gpq:.1f} MiB  ({tot_pct:.1f}%)  "
        f"elapsed: {_fmt_duration(totals['elapsed'])}"
    )
    print(sep)

# ---------------------------------------------------------------------------
# List command
# ---------------------------------------------------------------------------

def list_geoparquets() -> None:
    """Scan *dest_dir* and print a table of every GeoParquet with its CRS."""
    import pyarrow.parquet as pq

    if not _DEST_DIR.exists():
        print(f"  Output directory not found: {_DEST_DIR}")
        return

    files = sorted(_DEST_DIR.rglob("*.parquet"))
    if not files:
        print(f"  No GeoParquet files found in {_DEST_DIR}")
        return

    rows: list[dict] = []
    for p in files:
        try:
            meta = pq.ParquetFile(p).schema_arrow.metadata or {}
            geo  = json.loads(meta[b"geo"]) if b"geo" in meta else {}
            crs_entries = []
            for col_meta in geo.get("columns", {}).values():
                crs = col_meta.get("crs")
                if crs is None:
                    crs_entries.append("null")
                elif isinstance(crs, dict):
                    aid  = crs.get("id", {})
                    auth = aid.get("authority", "")
                    code = aid.get("code", "")
                    name = crs.get("name", "")
                    crs_entries.append(f"{auth}:{code}  ({name})" if auth and code else name or "unknown")
                else:
                    crs_entries.append(str(crs)[:60])
            crs_str = "; ".join(crs_entries) if crs_entries else "no geo metadata"
            mb = p.stat().st_size / 1_048_576
        except Exception as exc:
            crs_str = f"error: {exc}"
            mb = 0.0

        try:
            rel = str(p.relative_to(_DEST_DIR))
        except ValueError:
            rel = str(p)
        rows.append({"path": rel, "mb": mb, "crs": crs_str})

    WP  = min(max(len(r["path"]) for r in rows), 80)
    WC  = max(max(len(r["crs"])  for r in rows), 10)
    sep = "=" * (WP + 10 + WC + 6)

    print(f"\n{sep}")
    print(f"  {'File':<{WP}}  {'MB':>7}  CRS")
    print("-" * len(sep))
    for r in rows:
        print(f"  {r['path']:<{WP}}  {r['mb']:>7.1f}  {r['crs']}")
    print(sep)

    total_mb = sum(r["mb"] for r in rows)
    null_crs = sum(1 for r in rows if any(x in r["crs"] for x in ("null", "no geo", "error")))
    print(f"  Files: {len(rows)}  |  Total: {total_mb:.1f} MB  |  Missing CRS: {null_crs}")
    print(sep)

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    """Run the full PRODES → GeoParquet pipeline."""
    t0  = time.perf_counter()
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"\n{'─' * 60}")
    print(f"  PRODES → GeoParquet  v{__version__}  |  {now}")
    print(f"{'─' * 60}")

    convert_fn = _load_converter()

    if not _SOURCE_DIR.exists():
        sys.exit(f"[FATAL] source directory not found: {_SOURCE_DIR}")

    # 1. Existing GeoParquets
    existing: set[Path] = (
        {p.resolve() for p in _DEST_DIR.rglob("*.parquet") if p.stat().st_size > 0}
        if _DEST_DIR.exists() else set()
    )
    print(f"\n  [1/3] existing GeoParquets : {len(existing)}")

    # 2. Build manifest
    zip_files = sorted(_SOURCE_DIR.rglob("*.zip"))
    print(f"  [2/3] scanning {len(zip_files)} zip archive(s) ...")
    if not zip_files:
        sys.exit("  No zip archives found — nothing to do.")

    manifest = build_manifest(zip_files)
    n_gpkg   = sum(1 for j in manifest if j.kind == "gpkg")
    n_shp    = sum(1 for j in manifest if j.kind == "shp")
    print(
        f"        manifest: {len(manifest)} job(s)  "
        f"(gpkg layers: {n_gpkg}  orphan shapefiles: {n_shp})"
    )

    # 3. Cross-reference
    done_jobs = [j for j in manifest if j.out_path.resolve() in existing]
    todo_jobs = [j for j in manifest if j.out_path.resolve() not in existing]
    print(f"  [3/3] pending: {len(todo_jobs)}  |  already converted: {len(done_jobs)}")

    if not todo_jobs:
        print("\n  ✓ All files already converted.\n")
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
                "gpq_mb": _file_mb(j.out_path),
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
                pool.submit(_convert_job, job, convert_fn): (offset + i, job)
                for i, job in enumerate(todo_jobs, 1)
            }
            for fut in as_completed(futures):
                i, job              = futures[fut]
                status, src_mb, err = fut.result()
                biome  = job.rel_dir.replace("\\", "/").split("/")[0]
                label  = job.layer or Path(job.internal).stem
                gpq_mb = _file_mb(job.out_path) if status != "error" else None
                pct    = (gpq_mb / src_mb * 100) if (gpq_mb and src_mb) else None
                rows.append({
                    "n": i, "kind": job.kind, "biome": biome, "label": label,
                    "src_mb": src_mb, "gpq_mb": gpq_mb, "pct": pct, "status": status,
                })
                key = "ok" if status == "ok" else "errors" if status == "error" else "skipped"
                counts[key] += 1
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
                "version":  __version__,
                "config":   CONFIG,
                "summary":  {"total": len(manifest), **counts, "elapsed_s": round(elapsed, 1)},
                "jobs": [
                    {
                        "n":      r["n"],
                        "kind":   r["kind"],
                        "biome":  r["biome"],
                        "label":  r["label"],
                        "status": r["status"],
                        "src_mb": round(r["src_mb"], 2) if r["src_mb"] is not None else None,
                        "gpq_mb": round(r["gpq_mb"], 2) if r["gpq_mb"] is not None else None,
                        "pct":    round(r["pct"],    1)  if r["pct"]    is not None else None,
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
    import sys
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if cmd == "--list":
        list_geoparquets()
    else:
        main()
