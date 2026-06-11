"""
05_organize_geoparquet.py
=========================
Reorganizes the configured GeoParquet workspace by:

  1. MOVING (not copying) every parquet / tif file into a clean,
     date-agnostic folder structure:

       _organized\\
           <Biome>\\
               deforestation\\   â† accumulated / yearly deforestation layers
               auxiliary\\       â† borders, hydrography, land-use, etc.
               rasters\\         â† COG GeoTIFF files

     When the same filename exists in multiple dated sub-folders,
     only the most recently modified copy is kept; older duplicates
     are deleted.

  2. Removing the now-empty dated source directories.

  3. Writing README.md in the geoparquet root with the full catalog
     of every file (path, biome, type, size, row count).

After this script runs, ALL files live under _organized/.
The dated sub-folders (e.g. 2026-05-07/) are removed.

Usage
-----
    python 05_organize_geoparquet.py

Author
------
Amintas BrandÃ£o Jr. <abrandaojr@gmail.com>
Imazon â€” Instituto do Homem e Meio Ambiente da AmazÃ´nia

License
-------
MIT
"""

from __future__ import annotations

__version__ = "2.0.0"
__all__: list[str] = []

import importlib.util
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

HERE = Path(__file__).parent

# ---------------------------------------------------------------------------
# Dependency bootstrap
# ---------------------------------------------------------------------------

def _bootstrap(*packages: tuple[str, str]) -> None:
    import importlib, shutil as _sh
    mod_by_pip = {pip: mod for pip, mod in packages}

    def _still_missing(pkgs: list[str]) -> list[str]:
        importlib.invalidate_caches()
        return [p for p in pkgs if not importlib.util.find_spec(mod_by_pip[p])]

    missing = _still_missing(list(mod_by_pip))
    if not missing:
        return
    if not _sh.which("uv"):
        subprocess.call([sys.executable, "-m", "pip", "install", "--quiet", "uv"],
                        stderr=subprocess.DEVNULL)
    for base in [
        [sys.executable, "-m", "pip", "install", "--quiet"],
        ["uv", "pip", "install", "--python", sys.executable, "--quiet"],
        [sys.executable, "-m", "pip", "install", "--quiet", "--break-system-packages"],
    ]:
        if not missing:
            return
        try:
            subprocess.check_call(base + missing, stderr=subprocess.DEVNULL)
            missing = _still_missing(missing)
        except (subprocess.CalledProcessError, FileNotFoundError):
            pass
    if missing:
        sys.exit(f"[FATAL] Could not install: {' '.join(missing)}")


_bootstrap(("pyarrow", "pyarrow"), ("tqdm", "tqdm"))

from tqdm import tqdm   # noqa: E402

from prodes_pipeline.data_quality import (
    LineageRecord,
    StageTimer,
    atomic_write_text,
    configure_json_logging,
    file_inventory,
    freshness_metrics,
    parquet_quality_profile,
    require_existing_dir,
    to_jsonable,
    validate_nonempty_files,
    write_run_report,
)
from prodes_pipeline.pipeline_contracts import GEOPARQUET_CONTRACT
from prodes_pipeline.config import GEOPARQUET_DIR, REPORTS_DIR, ensure_pipeline_dirs

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

CONFIG: dict[str, object] = {
    "geoparquet_dir": GEOPARQUET_DIR,
}

SEP = "=" * 65
DIV = "-" * 65
REPORT_DIR = REPORTS_DIR
OBS_LOG = configure_json_logging(REPORT_DIR / "observability.jsonl")

# ---------------------------------------------------------------------------
# Classification rules
# ---------------------------------------------------------------------------

_BIOME_NAMES = {
    "Amazon Biome", "Legal Amazon", "Cerrado", "Caatinga",
    "Pantanal", "Mata Atlantica", "Pampa",
}

_DEFOR_KEYWORDS  = ("deforestation", "desmatamento", "desmat")
_RASTER_SUFFIXES = {".tif", ".tiff"}
_AUX_KEYWORDS    = (
    "border", "boundary", "hydrography", "hydro", "indigenous",
    "conservation", "settlement", "quilombola", "terra_indigena",
    "unidade_conservacao", "limite", "carbon", "biomass", "uc_",
)
_DATA_SUFFIXES = {".parquet", ".tif", ".tiff"}


def _classify(parts: tuple[str, ...]) -> str:
    path_low = "/".join(p.lower() for p in parts)
    if parts[-1].lower().endswith(tuple(_RASTER_SUFFIXES)):
        return "rasters"
    if any(k in path_low for k in _DEFOR_KEYWORDS):
        return "deforestation"
    return "auxiliary"


def _find_biome(parts: tuple[str, ...]) -> str | None:
    return next((p for p in parts if p in _BIOME_NAMES), None)

# ---------------------------------------------------------------------------
# File discovery: deduplicate, keep newest
# ---------------------------------------------------------------------------

def discover_files(root: Path) -> dict[tuple[str, str, str], Path]:
    """
    Scan root for data files outside _organized/.
    Returns {(biome, data_type, filename): best_path} where 'best' = newest mtime.
    """
    best: dict[tuple[str, str, str], tuple[float, Path]] = {}

    for f in root.rglob("*"):
        try:
            if (
                not f.is_file()
                or f.suffix.lower() not in _DATA_SUFFIXES
                or "_organized" in f.parts
                or f.name == "README.md"
                or "_catalog" in f.name
            ):
                continue
            rel   = f.relative_to(root)
            parts = rel.parts
            biome = _find_biome(parts) or "Unknown"
            dtype = _classify(parts)
            key   = (biome, dtype, f.name)
            st = f.stat()
            current = best.get(key)
            if current is None or st.st_mtime > current[0]:
                best[key] = (st.st_mtime, f)
        except (ValueError, OSError):
            pass

    return {key: path for key, (_, path) in best.items()}

# ---------------------------------------------------------------------------
# Move files into _organized/
# ---------------------------------------------------------------------------

def reorganize(root: Path, best: dict[tuple[str, str, str], Path]) -> Path:
    """
    Move winning files to _organized/<biome>/<data_type>/<filename>.
    Older duplicates (not in 'best') are deleted.
    Returns the _organized/ path.
    """
    organized = root / "_organized"

    total = len(best)
    print(f"\n  Moving {total} file(s) into _organized/ ...")
    bar = tqdm(best.items(), desc="  Moving", unit="file", ncols=80)

    for (biome, dtype, name), src in bar:
        bar.set_postfix_str(name[:40])
        biome_safe = biome.replace(" ", "_")
        dst = organized / biome_safe / dtype / name
        dst.parent.mkdir(parents=True, exist_ok=True)

        if dst == src:
            continue   # already in place

        if dst.exists():
            # Overwrite only if source is newer
            if src.stat().st_mtime > dst.stat().st_mtime:
                dst.unlink()
                shutil.move(str(src), dst)
            else:
                src.unlink()   # dst is newer â€” drop the older duplicate
        else:
            shutil.move(str(src), dst)

    bar.close()

    # Delete all remaining data files outside _organized/ (older duplicates)
    orphans = [
        p for p in root.rglob("*")
        if p.is_file()
        and p.suffix.lower() in _DATA_SUFFIXES
        and "_organized" not in p.parts
        and p.name != "README.md"
        and "_catalog" not in p.name
    ]
    if orphans:
        print(f"  Removing {len(orphans)} older duplicate(s)...")
        for p in orphans:
            try:
                p.unlink()
            except OSError as exc:
                print(f"  [WARN] could not delete {p.name}: {exc}")

    return organized


def remove_empty_dirs(root: Path) -> int:
    """
    Walk root bottom-up and remove empty directories (excluding _organized/).
    Returns count of directories removed.
    """
    removed = 0
    for d in sorted(root.rglob("*"), reverse=True):
        if not d.is_dir():
            continue
        if "_organized" in d.parts:
            continue
        if d == root:
            continue
        try:
            d.rmdir()   # only succeeds if truly empty
            removed += 1
        except OSError:
            pass
    return removed

# ---------------------------------------------------------------------------
# README builder
# ---------------------------------------------------------------------------

def write_readme(root: Path, organized: Path) -> Path:
    """
    Write README.md to root listing every file in _organized/.
    """
    import pyarrow.parquet as pq

    readme = root / "README.md"
    now    = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

    rows: list[dict] = []
    for f in sorted(organized.rglob("*")):
        if not f.is_file():
            continue
        if f.suffix.lower() not in _DATA_SUFFIXES:
            continue
        try:
            rel_org  = f.relative_to(organized)
            parts    = rel_org.parts
            biome    = parts[0].replace("_", " ") if parts else "?"
            dtype    = parts[1] if len(parts) > 1 else "?"
            size_mb  = f.stat().st_size / 1_048_576
            rows_n   = None
            cols_n   = None
            if f.suffix.lower() == ".parquet":
                try:
                    meta   = pq.read_metadata(str(f))
                    rows_n = meta.num_rows
                    cols_n = meta.num_columns
                except Exception:
                    pass
            rows.append({
                "path":    str(f.relative_to(root)).replace("\\", "/"),
                "biome":   biome,
                "type":    dtype,
                "size_mb": size_mb,
                "rows":    rows_n,
                "cols":    cols_n,
            })
        except (ValueError, OSError):
            pass

    # Group by biome
    from collections import defaultdict
    by_biome: dict[str, list[dict]] = defaultdict(list)
    for r in rows:
        by_biome[r["biome"]].append(r)

    lines: list[str] = [
        "# PRODES GeoParquet â€” File Catalog",
        "",
        f"Generated: {now}  ",
        f"Root: `_organized/`  ",
        f"Total files: {len(rows)}  ",
        f"Total size: {sum(r['size_mb'] for r in rows):.1f} MB",
        "",
        "---",
        "",
    ]

    for biome in sorted(by_biome):
        lines.append(f"## {biome}")
        lines.append("")
        lines.append("| Path | Type | Size (MB) | Rows | Columns |")
        lines.append("|------|------|----------:|-----:|--------:|")
        for r in sorted(by_biome[biome], key=lambda x: x["path"]):
            path_link = f"`{r['path']}`"
            rows_s    = f"{r['rows']:,}" if r["rows"] is not None else "â€”"
            cols_s    = str(r["cols"])   if r["cols"] is not None else "â€”"
            lines.append(
                f"| {path_link} | {r['type']} "
                f"| {r['size_mb']:.1f} | {rows_s} | {cols_s} |"
            )
        lines.append("")

    lines += [
        "---",
        "",
        "**Sources**",
        "",
        "- INPE/PRODES â€” Sistema de Monitoramento do Desmatamento na AmazÃ´nia Brasileira",
        "- TerraBrasilis â€” https://terrabrasilis.dpi.inpe.br",
        "",
        f"*Generated by `05_organize_geoparquet.py` v{__version__}*",
    ]

    atomic_write_text(readme, "\n".join(lines))
    return readme

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

def print_summary(organized: Path) -> None:
    from collections import Counter

    files = [
        p for p in organized.rglob("*")
        if p.is_file() and p.suffix.lower() in _DATA_SUFFIXES
    ]
    by_biome = Counter(p.relative_to(organized).parts[0] for p in files)
    by_type  = Counter(p.relative_to(organized).parts[1] for p in files
                       if len(p.relative_to(organized).parts) > 1)
    total_mb = sum(p.stat().st_size for p in files) / 1_048_576

    print(f"\n{SEP}")
    print(f"  ORGANIZED STRUCTURE")
    print(DIV)
    print(f"  Total files  : {len(files)}")
    print(f"  Total size   : {total_mb:.1f} MB")
    print()
    print("  By data type:")
    for t, n in sorted(by_type.items()):
        print(f"    {t:<20}  {n:>4}")
    print()
    print("  By biome:")
    for b, n in sorted(by_biome.items()):
        print(f"    {b.replace('_', ' '):<30}  {n:>4}")
    print(SEP)

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    ensure_pipeline_dirs()
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"\n{SEP}")
    print(f"  GeoParquet Organizer  v{__version__}  |  {now}")
    print(f"{SEP}\n")

    root = Path(str(CONFIG["geoparquet_dir"]))
    require_existing_dir(root, "GeoParquet")

    print(f"  Directory: {root}\n")

    # 1. Discover all data files, keep best (newest) per logical identity
    print("  [1/4] Discovering files...")
    discover_timer = StageTimer("05_discover_geoparquet_files")
    best = discover_files(root)
    print(f"         {len(best)} unique file(s) found")
    input_files = list(best.values())
    input_quality = {
        "contract": to_jsonable(GEOPARQUET_CONTRACT),
        "inventory": file_inventory(input_files),
        "freshness": freshness_metrics(input_files, GEOPARQUET_CONTRACT.freshness),
        "parquet_profile": parquet_quality_profile(
            [p for p in input_files if p.suffix.lower() == ".parquet"],
            GEOPARQUET_CONTRACT,
        ),
    }
    OBS_LOG.emit(
        "stage_metrics",
        **to_jsonable(
            discover_timer.finish(
                "ok",
                input_row_count=None,
                output_row_count=len(best),
                anomalies={
                    "schema": input_quality["parquet_profile"].get(
                        "schema_anomalies", []
                    ),
                    "freshness": input_quality["freshness"].get("stale", []),
                },
            )
        ),
    )

    if not best:
        sys.exit("  No data files found. Run script 02 first.")

    # 2. Move files into _organized/
    print("\n  [2/4] Reorganizing...")
    reorg_timer = StageTimer("05_reorganize_files")
    organized = reorganize(root, best)
    output_files = [
        p
        for p in organized.rglob("*")
        if p.is_file() and p.suffix.lower() in _DATA_SUFFIXES
    ]
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
        "stage_metrics",
        **to_jsonable(
            reorg_timer.finish(
                "ok",
                input_row_count=len(input_files),
                output_row_count=len(output_files),
            )
        ),
    )

    # 3. Remove now-empty dated directories
    print("\n  [3/4] Cleaning up empty directories...")
    removed = remove_empty_dirs(root)
    print(f"         {removed} empty director(ies) removed")

    # 4. Write README
    print("\n  [4/4] Writing README.md...")
    readme = write_readme(root, organized)
    artifacts = validate_nonempty_files([readme], "organizer catalog")
    print(f"         Saved: {readme}")

    # Summary
    print_summary(organized)
    report_path = write_run_report(
        REPORT_DIR,
        Path(__file__).name,
        {
            "status": "ok",
            "version": __version__,
            "geoparquet_dir": str(root),
            "organized_dir": str(organized),
            "unique_files": len(best),
            "empty_dirs_removed": removed,
            "input_quality": input_quality,
            "output_quality": output_quality,
            "artifacts": artifacts,
            "lineage": LineageRecord(
                stage_name="05_organize_geoparquet",
                upstream_sources=[str(root)],
                transformation="Move latest data files into _organized structure, remove older duplicates, and write a catalog.",
                downstream_outputs=[str(organized), str(readme)],
                contracts=[GEOPARQUET_CONTRACT.name],
            ),
        },
    )
    print(f"  Quality report: {report_path}")
    print(f"\n  Done. All files are now in:\n  {organized}\n{SEP}\n")


if __name__ == "__main__":
    main()

