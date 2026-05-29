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
import subprocess
import sys

# ---------------------------------------------------------------------------
# Dependency bootstrap
# ---------------------------------------------------------------------------

def _bootstrap(*packages: tuple[str, str]) -> None:
    """Install missing packages at runtime."""
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
    ("matplotlib", "matplotlib"),
    ("numpy",      "numpy"),
)

import matplotlib.pyplot as plt  # noqa: E402
import numpy as np               # noqa: E402

# ---------------------------------------------------------------------------
# CONFIG  ← the only section that needs to be edited
# ---------------------------------------------------------------------------

CONFIG: dict[str, object] = {
    "output_path": "amazon_deforestation_norad.png",
    "dpi":         220,
}

SEP = "=" * 65
DIV = "-" * 65

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

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    from datetime import datetime
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"\n{SEP}")
    print(f"  Deforestation Chart  v{__version__}  |  {now}")
    print(f"{SEP}\n")

    years  = sorted(ANNUAL_KM2) + sorted(TARGETS)
    values = [ANNUAL_KM2[y] for y in sorted(ANNUAL_KM2)] + [TARGETS[y] for y in sorted(TARGETS)]
    colors = [LGRAY] * len(ANNUAL_KM2) + [GREEN] * len(TARGETS)
    pos    = np.arange(len(years), dtype=float)

    plt.rcParams.update({
        "font.family":     "sans-serif",
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
        is_green  = yr in TARGETS
        txt_color = GREEN if is_green else DGRAY
        fw        = "bold" if yr in important else "normal"
        fs        = 10.5  if yr in important else 9.0
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
    ax.text(bx + 0.18, (TARGETS[2028] + ANNUAL_KM2[2025]) / 2, "−30%",
            va="center", fontsize=10.5, color=DGRAY, fontweight="bold")

    div_x = (pos[i25] + pos[i25 + 1]) / 2
    ax.axvline(div_x, color="#DDDDDD", lw=0.9, ls="--", zorder=2)
    ax.text(div_x - 0.15, 15600, "Observed",
            ha="right", va="top", fontsize=9, color="#BBBBBB", style="italic")
    ax.text(div_x + 0.15, 15600, "Projected",
            ha="left",  va="top", fontsize=9, color=GREEN,    style="italic")

    fig.text(0.0, 1.12,
             "Project Goal: Reduce annual deforestation in Brazil's Legal Amazon to 4,000 km² by 2028",
             fontsize=15, fontweight="bold", color="#111111",
             transform=ax.transAxes, va="top")
    fig.text(0.0, 1.058,
             "Annual forest loss in the Brazilian Amazon (sq. km). Green bars are project targets.",
             fontsize=10, color="#888888",
             transform=ax.transAxes, va="top", linespacing=1.5)
    fig.text(0.0, -0.06,
             "Source: INPE/PRODES. Historical peak: 29,059 km² in 1995.",
             fontsize=8.5, color="#BBBBBB",
             transform=ax.transAxes, va="top")

    plt.subplots_adjust(left=0.04, right=0.94, top=0.78, bottom=0.10)

    output = str(CONFIG["output_path"])
    plt.savefig(output, dpi=int(CONFIG["dpi"]), bbox_inches="tight", facecolor="white")

    print(f"  saved: {output}")
    print(f"\n{SEP}\n")


if __name__ == "__main__":
    main()
