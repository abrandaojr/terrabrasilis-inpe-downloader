"""
05_organize_geoparquet.py
=========================
Organizes C:\\Amintas\\Prodes\\geoparquet by:

  1. Scanning all parquet files and building a full catalog (_catalog.json).
  2. Creating a clean, date-agnostic folder structure (_organized\\) that
     groups files by biome and data type, always pointing to the most
     recently modified copy of each file.

The original files are NEVER moved or deleted.  _organized\\ contains
hard-copies of the winning (most recent) version of each logical file.

Output
------
  C:\\Amintas\\Prodes\\geoparquet\\_catalog.json
  C:\\Amintas\\Prodes\\geoparquet\\_organized\\
      <Biome>\\
          deforestation\\   ← accumulated / yearly deforestation layers
          auxiliary\\       ← borders, hydrography, land-use, etc.
          rasters\\         ← COG GeoTIFF files

Usage
-----
    python 05_organize_geoparquet.py

Author
------
Amintas Brandão Jr. <abrandaojr@gmail.com>
Imazon — Instituto do Homem e Meio Ambiente da Amazônia

License
-------
MIT
"""

from __future__ import annotations

__version__ = "1.0.0"
__all__: list[str] = []

import importlib.util
import json
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

HERE = Path(__file__).parent

# ---------------------------------------------------------------------------
# Dependency bootstrap
# ---------------------------------------------------------------------------

def _bootstrap(*packages: tuple[str, str]) -> None:
    import importlib, shutil as _shutil
    mod_by_pip = {pip: mod for pip, mod in packages}

    def _still_missing(pkgs: list[str]) -> list[str]:
        importlib.invalidate_caches()
        return [p for p in pkgs if not importlib.util.find_spec(mod_by_pip[p])]

    missing = _still_missing(list(mod_by_pip))
    if not missing:
        return
    if not _shutil.which("uv"):
        subprocess.call([sys.executable, "-m", "pip", "install", "--quiet", "uv"],
                        stderr=subprocess.DEVNULL)
    strategies = [
        [sys.executable, "-m", "pip", "install", "--quiet"],
        ["uv", "pip", "install", "--python", sys.executable, "--quiet"],
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


_bootstrap(("pyarrow", "pyarrow"), ("tqdm", "tqdm"))

from tqdm import tqdm   # noqa: E402

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

CONFIG: dict[str, object] = {
    "geoparquet_dir": r"C:\Amintas\Prodes\geoparquet",
}

SEP = "=" * 65
DIV = "-" * 65

# ---------------------------------------------------------------------------
# Classification rules
# ---------------------------------------------------------------------------

_BIOME_NAMES = {
    "Amazon Biome", "Legal Amazon", "Cerrado", "Caatinga",
    "Pantanal", "Mata Atlantica", "Pampa",
}

_DEFOR_KEYWORDS  = ("deforestation", "desmatamento", "desmat")
_RASTER_SUFFIXES = (".tif", ".tiff")
_AUX_KEYWORDS    = (
    "border", "boundary", "hydrography", "hydro", "indigenous",
    "conservation", "settlement", "quilombola", "terra_indigena",
    "unidade_conservacao", "limite", "carbon", "biomass", "uc_",
)


def _classify(parts: tuple[str, ...]) -> str:
    """Return 'deforestation' | 'raster' | 'auxiliary'."""
    path_low = "/".join(p.lower() for p in parts)
    if parts[-1].lower().endswith(_RASTER_SUFFIXES):
        return "raster"
    if any(k in path_low for k in _DEFOR_KEYWORDS):
        return "deforestation"
    if any(k in path_low for k in _AUX_KEYWORDS):
        return "auxiliary"
    return "auxiliary"   # default unknown → auxiliary


def _find_biome(parts: tuple[str, ...]) -> str | None:
    return next((p for p in parts if p in _BIOME_NAMES), None)


# ---------------------------------------------------------------------------
# Catalog builder
# ---------------------------------------------------------------------------

def build_catalog(root: Path) -> list[dict]:
    """Walk root recursively, build a metadata record for each data file."""
    import pyarrow.parquet as pq

    all_files = sorted(
        [p for p in root.rglob("*")
         if p.is_file() and p.suffix.lower() in (".parquet", ".tif", ".tiff")
         and "_organized" not in p.parts
         and "_catalog" not in p.name]
    )

    records: list[dict] = []
    bar = tqdm(all_files, desc="  Cataloging", unit="file", ncols=80)

    for f in bar:
        bar.set_postfix_str(f.name[:40])
        try:
            rel    = f.relative_to(root)
            parts  = rel.parts
            biome  = _find_biome(parts)
            dtype  = _classify(parts)
            size   = f.stat().st_size
            mtime  = f.stat().st_mtime

            schema_cols: list[str] = []
            row_count: int | None  = None
            if f.suffix.lower() == ".parquet":
                try:
                    meta = pq.read_metadata(str(f))
                    schema_cols = [meta.schema.column(i).name
                                   for i in range(meta.num_columns)]
                    row_count   = meta.num_rows
                except Exception:
                    pass

            records.append({
                "path":       str(rel).replace("\\", "/"),
                "filename":   f.name,
                "biome":      biome,
                "data_type":  dtype,
                "size_bytes": size,
                "mtime":      mtime,
                "columns":    schema_cols,
                "row_count":  row_count,
            })
        except Exception as exc:
            tqdm.write(f"  [WARN] {f.name}: {exc}")

    bar.close()
    return records


# ---------------------------------------------------------------------------
# Organized folder builder
# ---------------------------------------------------------------------------

def build_organized(root: Path, catalog: list[dict]) -> Path:
    """
    Create root/_organized/ with the most-recent copy of each logical file
    grouped as  <biome>/<data_type>/<filename>.
    When a filename appears in multiple dated folders, the newest wins.
    """
    organized = root / "_organized"

    # Deduplicate by (biome, data_type, filename) — keep newest mtime
    best: dict[tuple[str, str, str], dict] = {}
    for rec in catalog:
        biome = rec["biome"] or "Unknown"
        dtype = rec["data_type"]
        name  = rec["filename"]
        key   = (biome, dtype, name)
        if key not in best or rec["mtime"] > best[key]["mtime"]:
            best[key] = rec

    total = len(best)
    print(f"\n  Building _organized/ ({total} unique file(s))...")
    bar = tqdm(best.values(), desc="  Copying", unit="file", ncols=80)

    for rec in bar:
        bar.set_postfix_str(rec["filename"][:40])
        src  = root / rec["path"].replace("/", "\\")
        biome_safe = (rec["biome"] or "Unknown").replace(" ", "_")
        dst  = organized / biome_safe / rec["data_type"] / rec["filename"]
        dst.parent.mkdir(parents=True, exist_ok=True)
        try:
            if not dst.exists() or dst.stat().st_mtime < src.stat().st_mtime:
                shutil.copy2(src, dst)
        except Exception as exc:
            tqdm.write(f"  [WARN] copy failed: {rec['filename']}: {exc}")

    bar.close()
    return organized


# ---------------------------------------------------------------------------
# Summary report
# ---------------------------------------------------------------------------

def print_summary(catalog: list[dict]) -> None:
    from collections import Counter

    biome_counts  = Counter(r["biome"] or "Unknown" for r in catalog)
    dtype_counts  = Counter(r["data_type"]           for r in catalog)
    total_size_mb = sum(r["size_bytes"] for r in catalog) / 1_048_576

    print(f"\n{SEP}")
    print(f"  CATALOG SUMMARY")
    print(DIV)
    print(f"  Total files        : {len(catalog)}")
    print(f"  Total size         : {total_size_mb:.1f} MB")
    print()
    print(f"  By data type:")
    for dt, n in sorted(dtype_counts.items()):
        print(f"    {dt:<20}  {n:>4}")
    print()
    print(f"  By biome:")
    for bm, n in sorted(biome_counts.items()):
        print(f"    {bm:<30}  {n:>4}")
    print(SEP)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"\n{SEP}")
    print(f"  GeoParquet Organizer  v{__version__}  |  {now}")
    print(f"{SEP}\n")

    root = Path(str(CONFIG["geoparquet_dir"]))
    if not root.exists():
        sys.exit(f"[FATAL] GeoParquet directory not found: {root}\n"
                 "        Run  python 02_convert_to_geoparquet.py  first.")

    print(f"  Directory: {root}\n")

    # 1. Build catalog
    print("  [1/3] Building catalog...")
    catalog = build_catalog(root)

    catalog_path = root / "_catalog.json"
    with open(catalog_path, "w", encoding="utf-8") as f:
        json.dump(
            {
                "generated_at": datetime.now(timezone.utc).isoformat(),
                "total_files":  len(catalog),
                "files":        catalog,
            },
            f, ensure_ascii=False, indent=2, default=str,
        )
    print(f"  Catalog saved: {catalog_path}")

    # 2. Build organized view
    print("\n  [2/3] Creating organized view...")
    organized = build_organized(root, catalog)
    print(f"  Organized view: {organized}")

    # 3. Summary
    print("\n  [3/3] Summary:")
    print_summary(catalog)

    print(f"\n  Done.\n{SEP}\n")


if __name__ == "__main__":
    main()
