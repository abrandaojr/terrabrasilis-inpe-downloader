"""
05_organize_geoparquet.py
=========================
Reorganizes C:\\Amintas\\Prodes\\geoparquet by:

  1. MOVING (not copying) every parquet / tif file into a clean,
     date-agnostic folder structure:

       _organized\\
           <Biome>\\
               deforestation\\   ← accumulated / yearly deforestation layers
               auxiliary\\       ← borders, hydrography, land-use, etc.
               rasters\\         ← COG GeoTIFF files

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
    best: dict[tuple[str, str, str], Path] = {}

    all_files = [
        p for p in root.rglob("*")
        if p.is_file()
        and p.suffix.lower() in _DATA_SUFFIXES
        and "_organized" not in p.parts
        and p.name != "README.md"
        and "_catalog" not in p.name
    ]

    for f in all_files:
        try:
            rel   = f.relative_to(root)
            parts = rel.parts
            biome = _find_biome(parts) or "Unknown"
            dtype = _classify(parts)
            key   = (biome, dtype, f.name)
            if key not in best or f.stat().st_mtime > best[key].stat().st_mtime:
                best[key] = f
        except (ValueError, OSError):
            pass

    return best

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

    # Set of canonical destinations
    destinations: set[Path] = set()
    for (biome, dtype, name), src in best.items():
        biome_safe = biome.replace(" ", "_")
        dst = organized / biome_safe / dtype / name
        destinations.add(dst)

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
                src.unlink()   # dst is newer — drop the older duplicate
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
        "# PRODES GeoParquet — File Catalog",
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
            rows_s    = f"{r['rows']:,}" if r["rows"] is not None else "—"
            cols_s    = str(r["cols"])   if r["cols"] is not None else "—"
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
        "- INPE/PRODES — Sistema de Monitoramento do Desmatamento na Amazônia Brasileira",
        "- TerraBrasilis — https://terrabrasilis.dpi.inpe.br",
        "",
        f"*Generated by `05_organize_geoparquet.py` v{__version__}*",
    ]

    readme.write_text("\n".join(lines), encoding="utf-8")
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
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"\n{SEP}")
    print(f"  GeoParquet Organizer  v{__version__}  |  {now}")
    print(f"{SEP}\n")

    root = Path(str(CONFIG["geoparquet_dir"]))
    if not root.exists():
        sys.exit(f"[FATAL] GeoParquet directory not found: {root}\n"
                 "        Run  python 02_convert_to_geoparquet.py  first.")

    print(f"  Directory: {root}\n")

    # 1. Discover all data files, keep best (newest) per logical identity
    print("  [1/4] Discovering files...")
    best = discover_files(root)
    print(f"         {len(best)} unique file(s) found")

    if not best:
        sys.exit("  No data files found. Run script 02 first.")

    # 2. Move files into _organized/
    print("\n  [2/4] Reorganizing...")
    organized = reorganize(root, best)

    # 3. Remove now-empty dated directories
    print("\n  [3/4] Cleaning up empty directories...")
    removed = remove_empty_dirs(root)
    print(f"         {removed} empty director(ies) removed")

    # 4. Write README
    print("\n  [4/4] Writing README.md...")
    readme = write_readme(root, organized)
    print(f"         Saved: {readme}")

    # Summary
    print_summary(organized)
    print(f"\n  Done. All files are now in:\n  {organized}\n{SEP}\n")


if __name__ == "__main__":
    main()
