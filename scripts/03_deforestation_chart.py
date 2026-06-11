"""
03_deforestation_chart.py
=========================
Generates a publication-ready bar chart of annual Amazon deforestation
rates (INPE/PRODES), including historical data and project targets.

Output
------
Saves the chart as a PNG file (path configured in CONFIG).

Usage
-----
    python 03_deforestation_chart.py

Author
------
Amintas BrandÃ£o Jr. <abrandaojr@gmail.com>
Imazon â€” Instituto do Homem e Meio Ambiente da AmazÃ´nia

License
-------
MIT
"""

from __future__ import annotations

__version__ = "1.0.0"

# ---------------------------------------------------------------------------
# Standard library imports
# ---------------------------------------------------------------------------
import importlib
import importlib.util
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

# ---------------------------------------------------------------------------
# Local constants
# ---------------------------------------------------------------------------
HERE = Path(__file__).parent


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
    mod_by_pip = {pip: mod for pip, mod in packages}

    def _still_missing(pkgs: list[str]) -> list[str]:
        importlib.invalidate_caches()
        return [p for p in pkgs if not importlib.util.find_spec(mod_by_pip[p])]

    missing = _still_missing(list(mod_by_pip))
    if not missing:
        return

    # Ensure uv is available if possible, as it's a preferred strategy
    if not shutil.which("uv"):
        try:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", "--quiet", "uv"],
                stderr=subprocess.DEVNULL,
            )
            importlib.invalidate_caches()  # Invalidate after uv install
        except subprocess.CalledProcessError:
            print("[WARNING] Failed to install 'uv'. Falling back to pip for dependencies.", file=sys.stderr)
        except FileNotFoundError:
            print("[WARNING] Python executable or pip not found to install 'uv'. Falling back to other strategies.", file=sys.stderr)

    strategies = [
        [sys.executable, "-m", "pip", "install", "--quiet"],
    ]
    if shutil.which("uv"):
        strategies.append(["uv", "pip", "install", "--python", sys.executable, "--quiet"])
    # Fallback to `python -m uv` if `uv` executable is not on PATH but module is available
    elif importlib.util.find_spec("uv"):
        strategies.append([sys.executable, "-m", "uv", "pip", "install", "--python", sys.executable, "--quiet"])

    strategies.append([sys.executable, "-m", "pip", "install", "--quiet", "--break-system-packages"])

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


# Call bootstrap to ensure dependencies are installed BEFORE importing them
_bootstrap(
    ("matplotlib", "matplotlib"),
    ("numpy", "numpy"),
)


# ---------------------------------------------------------------------------
# Third-party library imports (now safe to import)
# ---------------------------------------------------------------------------
import matplotlib.pyplot as plt
import numpy as np

from prodes_pipeline.data_quality import (
    LineageRecord,
    StageTimer,
    configure_json_logging,
    numeric_distribution,
    to_jsonable,
    validate_nonempty_files,
    write_run_report,
)
from prodes_pipeline.config import FIGURES_DIR, REPORTS_DIR, ensure_pipeline_dirs


# ---------------------------------------------------------------------------
# CONFIG  â† the only section that needs to be edited
# ---------------------------------------------------------------------------

CONFIG: dict[str, object] = {
    "output_path": FIGURES_DIR / "amazon_deforestation_norad.png",
    "dpi": 220,
}

SEP = "=" * 65
REPORT_DIR = REPORTS_DIR
OBS_LOG = configure_json_logging(REPORT_DIR / "observability.jsonl")
# DIV is defined but not used. Removed for PEP 8.

# ---------------------------------------------------------------------------
# Data
# ---------------------------------------------------------------------------

ANNUAL_KM2: dict[int, int] = {
    2015: 6207, 2016: 7893, 2017: 6947, 2018: 7536,
    2019: 10129, 2020: 10851, 2021: 13038,
    2022: 11594, 2023: 9064, 2024: 6518,
    2025: 5731,
}
TARGETS: dict[int, int] = {2026: 4866, 2028: 4000}

GREEN = "#2E7D32"
LGRAY = "#CCCCCC"
DGRAY = "#444444"


def _validate_inputs() -> dict[str, object]:
    years = sorted(ANNUAL_KM2)
    target_years = sorted(TARGETS)
    if len(years) != len(ANNUAL_KM2):
        raise SystemExit("[FATAL] Duplicate annual deforestation years detected.")
    if any(v <= 0 for v in ANNUAL_KM2.values()):
        raise SystemExit("[FATAL] Annual deforestation values must be positive.")
    if any(v <= 0 for v in TARGETS.values()):
        raise SystemExit("[FATAL] Target values must be positive.")
    if target_years and years and min(target_years) <= max(years):
        raise SystemExit("[FATAL] Target years must be after observed years.")
    return {
        "observed_year_min": min(years),
        "observed_year_max": max(years),
        "observed_count": len(years),
        "target_years": target_years,
    }


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> None:
    """Generates and saves the Amazon deforestation chart."""
    ensure_pipeline_dirs()
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"\n{SEP}")
    print(f"  Deforestation Chart  v{__version__}  |  {now}")
    print(f"{SEP}\n")

    input_quality = _validate_inputs()
    input_quality["annual_distribution"] = numeric_distribution(
        list(ANNUAL_KM2.values())
    )
    stage_timer = StageTimer("03_generate_deforestation_chart")
    years = sorted(ANNUAL_KM2) + sorted(TARGETS)
    values = ([ANNUAL_KM2[y] for y in sorted(ANNUAL_KM2)] +
              [TARGETS[y] for y in sorted(TARGETS)])
    colors = [LGRAY] * len(ANNUAL_KM2) + [GREEN] * len(TARGETS)
    pos = np.arange(len(years), dtype=float)

    plt.rcParams.update({
        "font.family": "sans-serif",
        "font.sans-serif": ["Arial", "Helvetica Neue", "DejaVu Sans"],
    })

    fig, ax = plt.subplots(figsize=(12, 6.5))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    for sp in ax.spines.values():
        sp.set_visible(False)

    ax.bar(pos, values, width=0.65, color=colors, zorder=3, linewidth=0)
    ax.yaxis.grid(True, color="#F2F2F2", linewidth=0.8, zorder=0)
    ax.set_axisbelow(True)
    ax.axhline(16000, color="#111", lw=2.0, zorder=6, clip_on=False)

    ax.set_xticks(pos)
    ax.set_xticklabels([str(y) for y in years], fontsize=10, color="#888888")
    ax.xaxis.set_tick_params(length=0)
    ax.set_ylim(0, 16200)
    ax.yaxis.set_visible(False)

    important = {2021, 2025, 2028}
    for i, (yr, val) in enumerate(zip(years, values)):
        is_green = yr in TARGETS
        txt_color = GREEN if is_green else DGRAY
        fw = "bold" if yr in important else "normal"
        fs = 10.5 if yr in important else 9.0
        ax.text(pos[i], val + 220, f"{val:,}",
                ha="center", va="bottom",
                fontsize=fs, color=txt_color, fontweight=fw)

    i25 = years.index(2025)
    i28 = years.index(2028)

    ax.text(pos[i25], ANNUAL_KM2[2025] + 220 + 900, "Baseline",
            ha="center", va="bottom", fontsize=8.5, color=DGRAY, style="italic")
    ax.text(pos[i28], TARGETS[2028] + 220 + 900, "Target",
            ha="center", va="bottom", fontsize=8.5, color=GREEN, style="italic")

    ax.hlines(ANNUAL_KM2[2025], pos[i25], pos[i28] + 0.32,
              colors="#BBBBBB", lw=1.0, ls=":", zorder=4)

    bx = pos[i28] + 0.52
    ax.annotate("", xy=(bx, TARGETS[2028]), xytext=(bx, ANNUAL_KM2[2025]),
                arrowprops=dict(arrowstyle="<->", color=DGRAY, lw=1.1))
    ax.text(bx + 0.18, (TARGETS[2028] + ANNUAL_KM2[2025]) / 2, "âˆ’30%",
            va="center", fontsize=10.5, color=DGRAY, fontweight="bold")

    div_x = (pos[i25] + pos[i25 + 1]) / 2
    ax.axvline(div_x, color="#DDDDDD", lw=0.9, ls="--", zorder=2)
    ax.text(div_x - 0.15, 15600, "Observed",
            ha="right", va="top", fontsize=9, color="#BBBBBB", style="italic")
    ax.text(div_x + 0.15, 15600, "Projected",
            ha="left", va="top", fontsize=9, color=GREEN, style="italic")

    fig.text(0.0, 1.12,
             ("Project Goal: Reduce annual deforestation in Brazil's Legal Amazon to "
              "4,000 kmÂ² by 2028"),
             fontsize=15, fontweight="bold", color="#111111",
             transform=ax.transAxes, va="top")
    fig.text(0.0, 1.058,
             ("Annual forest loss in the Brazilian Amazon (sq. km). "
              "Green bars are project targets."),
             fontsize=10, color="#888888",
             transform=ax.transAxes, va="top", linespacing=1.5)
    fig.text(0.0, -0.06,
             "Source: INPE/PRODES. Historical peak: 29,059 kmÂ² in 1995.",
             fontsize=8.5, color="#BBBBBB",
             transform=ax.transAxes, va="top")

    plt.subplots_adjust(left=0.04, right=0.94, top=0.78, bottom=0.10)

    output_path = Path(str(CONFIG["output_path"]))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output = str(output_path)
    plt.savefig(output, dpi=CONFIG["dpi"], bbox_inches="tight", facecolor="white")
    artifacts = validate_nonempty_files([Path(output)], "chart")
    metrics = stage_timer.finish(
        "ok",
        input_row_count=len(ANNUAL_KM2),
        output_row_count=len(artifacts),
    )
    OBS_LOG.emit("stage_metrics", **to_jsonable(metrics))
    report_path = write_run_report(
        REPORT_DIR,
        Path(__file__).name,
        {
            "status": "ok",
            "version": __version__,
            "input_quality": input_quality,
            "artifacts": artifacts,
            "lineage": LineageRecord(
                stage_name="03_deforestation_chart",
                upstream_sources=["ANNUAL_KM2 constant", "TARGETS constant"],
                transformation="Render publication PNG from observed annual PRODES values and project targets.",
                downstream_outputs=[output],
                contracts=[],
            ),
        },
    )

    print(f"  saved: {output}")
    print(f"  Quality report: {report_path}")
    print(f"{SEP}\n")


if __name__ == "__main__":
    main()

