"""
04_generate_presentation.py
===========================
Generates a bilingual press PowerPoint on Brazilian deforestation data (PRODES/INPE).

Structure
---------
Slides 1–10  : PT-BR (Português Brasileiro)
Slides 11–20 : EN-US (American English)

Each section covers:
  1  Cover
  2  Lead stat  (56 % drop)
  3  Amazon historical series  (2015–2028)
  4  By biome  (horizontal bar, 2023)
  5  Cerrado spotlight  (stat cards)
  6  Forest cover remaining  (% original, MapBiomas)
  7  Target trajectory  (bars + dashed line to 2028)
  8  International context  (GFW 2023)
  9  What explains the decline  (text + icons)
  10 Key takeaways  (3 cards)

Usage
-----
    python 04_generate_presentation.py

Output
------
    PRODES_Press_Briefing.pptx  (in the script directory)

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
import io
import subprocess
import sys
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Dependency bootstrap
# ---------------------------------------------------------------------------

def _bootstrap(*packages: tuple[str, str]) -> None:
    """Install missing packages into the current Python environment."""
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
    ("matplotlib", "matplotlib"),
    ("numpy",      "numpy"),
    ("python-pptx", "pptx"),
)

from datetime import datetime                      # noqa: E402 (already imported above)
import matplotlib.pyplot as plt                    # noqa: E402
import numpy as np                                 # noqa: E402
from pptx import Presentation                      # noqa: E402
from pptx.util import Inches, Pt                   # noqa: E402
from pptx.dml.color import RGBColor                # noqa: E402

# ---------------------------------------------------------------------------
# CONFIG  ← the only section that needs to be edited
# ---------------------------------------------------------------------------

CONFIG: dict[str, object] = {
    "output_path": "PRODES_Press_Briefing.pptx",
    "chart_dpi":   220,
}

SEP = "=" * 65
DIV = "-" * 65

# ---------------------------------------------------------------------------
# DATA  (curated, verified against PRODES/INPE and MapBiomas reports)
# ---------------------------------------------------------------------------

# Amazon Legal annual deforestation (km²) — INPE/PRODES
AMAZON_KM2: dict[int, int] = {
    2015: 6_207, 2016: 7_893, 2017: 6_947, 2018: 7_536,
    2019: 10_129, 2020: 10_851, 2021: 13_038,
    2022: 11_594, 2023: 9_064, 2024: 6_518, 2025: 5_731,
}
AMAZON_TARGETS: dict[int, int] = {2026: 4_866, 2028: 4_000}
HIST_PEAK_YEAR, HIST_PEAK_KM2 = 1995, 29_059

# Deforestation by biome — PRODES 2023 (km²)
BIOMES_KM2: dict[str, dict[str, int]] = {
    "pt": {
        "Amazônia Legal":  9_064,
        "Cerrado":         7_340,
        "Caatinga":        1_126,
        "Pantanal":          592,
        "Mata Atlântica":    103,
        "Pampa":              97,
    },
    "en": {
        "Legal Amazon":    9_064,
        "Cerrado Savanna": 7_340,
        "Caatinga":        1_126,
        "Pantanal":          592,
        "Atlantic Forest":   103,
        "Pampa":              97,
    },
}

# Remaining native vegetation (% of original biome area) — MapBiomas 2023
COVER_PCT: dict[str, dict[str, float]] = {
    "pt": {
        "Pantanal":        86.0,
        "Amazônia":        81.0,
        "Caatinga":        60.0,
        "Cerrado":         53.0,
        "Pampa":           38.0,
        "Mata Atlântica":  13.0,
    },
    "en": {
        "Pantanal":         86.0,
        "Amazon":           81.0,
        "Caatinga":         60.0,
        "Cerrado Savanna":  53.0,
        "Pampa":            38.0,
        "Atlantic Forest":  13.0,
    },
}

# Tropical primary forest loss 2023 (km²) — Global Forest Watch / FAO
# Note: GFW measures tree-cover loss (broader than PRODES deforestation).
# Brazil figure uses PRODES for comparability with national data.
INTL_LOSS_KM2: dict[str, int] = {
    "Brazil":      9_064,
    "D.R. Congo":  4_900,
    "Bolivia":     4_200,
    "Indonesia":   2_800,
    "Colombia":    1_450,
}

# ---------------------------------------------------------------------------
# STYLE
# ---------------------------------------------------------------------------

# Palette
C_GREEN  = "#2E7D32"
C_RED    = "#C0392B"
C_BLUE   = "#1A5276"
C_ORANGE = "#E67E22"
C_GRAY   = "#CCCCCC"
C_DARK   = "#111111"
C_MED    = "#555555"
C_LIGHT  = "#999999"

# Language accent colors
_LANG_COLOR = {"pt": "#1B5E20", "en": "#0D47A1"}

# Slide dimensions (16:9 widescreen)
_SW = 10.0    # width  in inches
_SH = 5.625   # height in inches


def _in(x: float) -> int:
    """Convert inches to EMU (English Metric Units)."""
    return int(x * 914_400)


def _rgb(hex_color: str) -> RGBColor:
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# ---------------------------------------------------------------------------
# BILINGUAL COPY
# ---------------------------------------------------------------------------

COPY: dict[str, dict[str, object]] = {
    # ── Sections ───────────────────────────────────────────────────────────
    "sec_amazon":   {"pt": "AMAZÔNIA LEGAL",      "en": "LEGAL AMAZON"},
    "sec_biomes":   {"pt": "POR BIOMA · 2023",    "en": "BY BIOME · 2023"},
    "sec_cerrado":  {"pt": "CERRADO EM FOCO",      "en": "CERRADO SPOTLIGHT"},
    "sec_cover":    {"pt": "COBERTURA FLORESTAL",  "en": "FOREST COVER"},
    "sec_target":   {"pt": "META 2028",            "en": "2028 TARGET"},
    "sec_global":   {"pt": "CONTEXTO GLOBAL",      "en": "GLOBAL CONTEXT"},
    "sec_causes":   {"pt": "CAUSAS & RISCOS",      "en": "DRIVERS & RISKS"},
    "sec_takeaway": {"pt": "MENSAGENS-CHAVE",      "en": "KEY TAKEAWAYS"},

    # ── Cover ──────────────────────────────────────────────────────────────
    "cover_title": {
        "pt": "Desmatamento no Brasil\nCaiu — Mas Ainda Não Chega.",
        "en": "Brazil's Deforestation\nIs Down — But Not Enough.",
    },
    "cover_sub": {
        "pt": "Dados PRODES · INPE  |  Apresentação para a Imprensa  ·  2026",
        "en": "PRODES · INPE Data  |  Press Briefing  ·  2026",
    },
    "cover_credit": {
        "pt": "Imazon — Instituto do Homem e Meio Ambiente da Amazônia",
        "en": "Imazon — Institute for People and the Environment in Amazonia",
    },

    # ── Slide 2: lead stat ─────────────────────────────────────────────────
    "stat_headline": {
        "pt": "O desmatamento na Amazônia Legal caiu 56% em quatro anos.",
        "en": "Amazon deforestation fell 56% in four years.",
    },
    "stat_body": {
        "pt": (
            "De 13.038 km² em 2021 para 5.731 km² em 2025 — "
            "a menor área desmatada desde 2015.\n"
            "Ainda assim, o dobro da meta estabelecida para 2028."
        ),
        "en": (
            "From 13,038 km² in 2021 to 5,731 km² in 2025 — "
            "the lowest figure since 2015.\n"
            "Yet still twice the target set for 2028."
        ),
    },
    "stat_ctx": {
        "pt": f"Pico histórico: {HIST_PEAK_KM2:,} km² em {HIST_PEAK_YEAR}".replace(",", "."),
        "en": f"Historical peak: {HIST_PEAK_KM2:,} km² in {HIST_PEAK_YEAR}",
    },

    # ── Slide 3: historical ────────────────────────────────────────────────
    "hist_headline": {
        "pt": "Pico em 2021, queda constante desde então — mas a meta ainda está longe.",
        "en": "Peak in 2021, steady decline since — but the 2028 target remains out of reach.",
    },
    "hist_sub": {
        "pt": "Desmatamento anual na Amazônia Legal (km²)  ·  2015–2028",
        "en": "Annual deforestation in Brazil's Legal Amazon (km²)  ·  2015–2028",
    },

    # ── Slide 4: by biome ──────────────────────────────────────────────────
    "biome_headline": {
        "pt": "O Cerrado perde mais vegetação que a Amazônia — e recebe menos atenção.",
        "en": "The Cerrado loses more native vegetation than the Amazon — and gets far less attention.",
    },
    "biome_sub": {
        "pt": "Desmatamento por bioma brasileiro (km²)  ·  PRODES 2023",
        "en": "Deforestation by Brazilian biome (km²)  ·  PRODES 2023",
    },

    # ── Slide 5: Cerrado spotlight ─────────────────────────────────────────
    "cerrado_headline": {
        "pt": "O Cerrado é o bioma brasileiro mais ameaçado proporcionalmente.",
        "en": "The Cerrado is Brazil's most proportionally threatened biome.",
    },
    "cerrado_card1_num":   {"pt": "7.340", "en": "7,340"},
    "cerrado_card1_label": {
        "pt": "km² desmatados no Cerrado (2023)\n— quase o mesmo que a Amazônia",
        "en": "km² deforested in the Cerrado (2023)\n— almost as much as the Amazon",
    },
    "cerrado_card2_num":   {"pt": "53%",   "en": "53%"},
    "cerrado_card2_label": {
        "pt": "de cobertura original remanescente\n— Mata Atlântica perdeu 87%",
        "en": "of original cover remaining\n— Atlantic Forest already lost 87%",
    },
    "cerrado_note": {
        "pt": (
            "O Cerrado abriga 5% da biodiversidade mundial e regula o ciclo hídrico "
            "das principais bacias hidrográficas do Brasil. "
            "Recebe, no entanto, menos de 10% dos recursos do Fundo Amazônia."
        ),
        "en": (
            "The Cerrado harbors 5% of the world's biodiversity and regulates "
            "the water cycle of Brazil's main river basins. "
            "Yet it receives less than 10% of Amazon Fund resources."
        ),
    },

    # ── Slide 6: forest cover ──────────────────────────────────────────────
    "cover_headline": {
        "pt": "Mata Atlântica: 13% restam. Pantanal: 86%. Cerrado: 53%.",
        "en": "Atlantic Forest: 13% remains. Pantanal: 86%. Cerrado: 53%.",
    },
    "cover_sub": {
        "pt": "Vegetação nativa remanescente por bioma (% da área original)  ·  MapBiomas 2023",
        "en": "Remaining native vegetation by biome (% of original area)  ·  MapBiomas 2023",
    },

    # ── Slide 7: target ────────────────────────────────────────────────────
    "target_headline": {
        "pt": "A meta de 4.000 km² para 2028 é alcançável — mas exige aceleração imediata.",
        "en": "The 4,000 km² target for 2028 is achievable — but requires immediate acceleration.",
    },
    "target_sub": {
        "pt": "Trajetória observada e meta de redução (km²)  ·  2015–2028",
        "en": "Observed trend and reduction target (km²)  ·  2015–2028",
    },

    # ── Slide 8: international ─────────────────────────────────────────────
    "intl_headline": {
        "pt": "Brasil lidera a queda global — mas ainda é o maior desmatador tropical.",
        "en": "Brazil leads the global decline — but remains the world's largest tropical deforester.",
    },
    "intl_sub": {
        "pt": "Perda de floresta tropical por país (km²)  ·  2023  ·  Fonte: GFW / FAO",
        "en": "Tropical forest loss by country (km²)  ·  2023  ·  Source: GFW / FAO",
    },
    "intl_note": {
        "pt": "* Metodologias diferentes. GFW mede perda de cobertura arbórea; PRODES mede desmatamento.",
        "en": "* Different methodologies. GFW measures tree-cover loss; PRODES measures deforestation.",
    },

    # ── Slide 9: causes ────────────────────────────────────────────────────
    "causes_headline": {
        "pt": "O que explica a queda — e o que pode revertê-la.",
        "en": "What drove the decline — and what could reverse it.",
    },
    "causes_items": {
        "pt": [
            ("▲", C_GREEN,  "Fiscalização reforçada",
             "Ibama e PF multiplicaram autuações e embargos desde 2023. "
             "Operações coordenadas reduziram o desmatamento ilegal."),
            ("▲", C_GREEN,  "Financiamento climático",
             "Fundo Amazônia recebeu +R$ 3 bi em 2023–24 (Noruega, Alemanha, EUA). "
             "Primeira fase do REDD+ Amazon operacional."),
            ("▼", C_RED,    "Risco: anistia fundiária",
             "Projetos de lei que regularizam desmatamentos ilegais até 2008 "
             "ameaçam criar incentivo para novos crimes ambientais."),
        ],
        "en": [
            ("▲", C_GREEN,  "Enforcement strengthened",
             "Ibama and Federal Police multiplied fines and embargoes from 2023. "
             "Coordinated operations reduced illegal deforestation."),
            ("▲", C_GREEN,  "Climate finance increased",
             "Amazon Fund received BRL 3 bn+ in 2023–24 (Norway, Germany, USA). "
             "First phase of REDD+ Amazon operational."),
            ("▼", C_RED,    "Risk: land regularization bills",
             "Legislation that would amnesty illegal deforestation before 2008 "
             "risks creating incentives for new environmental crimes."),
        ],
    },

    # ── Slide 10: takeaways ────────────────────────────────────────────────
    "takeaway_headline": {
        "pt": "Três mensagens desta apresentação.",
        "en": "Three messages from this briefing.",
    },
    "takeaways": {
        "pt": [
            ("A queda é real — mas frágil.",
             "56% de redução em 4 anos é histórico. "
             "Qualquer mudança na política de fiscalização pode revertê-la rapidamente."),
            ("O Cerrado está em crise silenciosa.",
             "Perde tanto quanto a Amazônia, tem 53% de cobertura original "
             "e recebe proporcionalmente muito menos recursos e atenção."),
            ("A meta de 2028 exige ação agora.",
             "Para alcançar 4.000 km², o Brasil precisa reduzir "
             "o desmatamento em mais 30% nos próximos três anos."),
        ],
        "en": [
            ("The decline is real — but fragile.",
             "A 56% drop in 4 years is historic. "
             "Any shift in enforcement policy could quickly reverse those gains."),
            ("The Cerrado is in silent crisis.",
             "It loses as much as the Amazon, retains only 53% of original cover, "
             "and receives far less resources and political attention."),
            ("The 2028 target requires action now.",
             "To reach 4,000 km², Brazil must cut deforestation "
             "by another 30% in the next three years."),
        ],
    },

    # ── Sources ────────────────────────────────────────────────────────────
    "src_prodes": {
        "pt": "Fonte: INPE/PRODES · Sistema de Monitoramento do Desmatamento na Amazônia Brasileira",
        "en": "Source: INPE/PRODES · Brazilian Amazon Deforestation Monitoring System",
    },
    "src_mapbiomas": {
        "pt": "Fonte: MapBiomas 2023 · Projeto de Mapeamento Anual da Cobertura e Uso da Terra",
        "en": "Source: MapBiomas 2023 · Annual Land Cover and Use Mapping Project",
    },
    "src_gfw": {
        "pt": "Fonte: Global Forest Watch / FAO 2023 · Dados aproximados — metodologias distintas",
        "en": "Source: Global Forest Watch / FAO 2023 · Approximate data — methodologies differ",
    },
}

# ---------------------------------------------------------------------------
# MATPLOTLIB CHART ENGINE
# ---------------------------------------------------------------------------

def _style() -> None:
    """Apply NYT/Economist-inspired rcParams."""
    plt.rcParams.update({
        "font.family":      "sans-serif",
        "font.sans-serif":  ["Arial", "Helvetica Neue", "Liberation Sans", "DejaVu Sans"],
        "figure.facecolor": "white",
        "axes.facecolor":   "white",
        "axes.grid":        True,
        "axes.grid.axis":   "y",
        "grid.color":       "#F0F0F0",
        "grid.linewidth":   0.8,
        "axes.spines.top":     False,
        "axes.spines.right":   False,
        "axes.spines.left":    False,
        "axes.spines.bottom":  True,
        "axes.axisbelow":      True,
        "xtick.bottom": False,
        "ytick.left":   False,
        "axes.labelcolor": C_MED,
        "xtick.color":     C_LIGHT,
    })


def _fig(w: float = 9.0, h: float = 3.8) -> tuple:
    _style()
    fig, ax = plt.subplots(figsize=(w, h))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    ax.spines["bottom"].set_color("#E0E0E0")
    ax.yaxis.set_visible(False)
    ax.xaxis.set_tick_params(length=0)
    return fig, ax


def _buf(fig) -> io.BytesIO:
    b = io.BytesIO()
    fig.savefig(b, format="png", dpi=int(CONFIG["chart_dpi"]),
                bbox_inches="tight", facecolor="white")
    b.seek(0)
    plt.close(fig)
    return b


def _label(ax, x, y, text, color=C_MED, size=8.5, bold=False, ha="center", va="bottom"):
    ax.text(x, y, text, ha=ha, va=va, fontsize=size, color=color,
            fontweight="bold" if bold else "normal")


# ── Chart 1: Amazon historical bar chart ────────────────────────────────────

def chart_amazon_historical(lang: str) -> io.BytesIO:
    years  = sorted(AMAZON_KM2) + [2026, 2028]
    values = [AMAZON_KM2[y] for y in sorted(AMAZON_KM2)] + [4_866, 4_000]
    colors = []
    for y in sorted(AMAZON_KM2):
        if y == 2021:
            colors.append(C_RED)
        elif y == 2025:
            colors.append(C_BLUE)
        else:
            colors.append(C_GRAY)
    colors += [C_GREEN, C_GREEN]

    pos = np.arange(len(years), dtype=float)
    fig, ax = _fig(9.2, 3.9)
    ax.bar(pos, values, width=0.72, color=colors, zorder=3, linewidth=0)
    ax.set_ylim(0, 15_200)
    ax.set_xticks(pos)
    ax.set_xticklabels([str(y) for y in years], fontsize=9, color=C_MED)

    for i, (y, v) in enumerate(zip(years, values)):
        is_peak = y == 2021
        is_curr = y == 2025
        is_tgt  = y in AMAZON_TARGETS
        c  = C_RED if is_peak else (C_GREEN if is_tgt else (C_BLUE if is_curr else C_MED))
        fs = 9.5 if (is_peak or is_curr) else 8.5
        _label(ax, pos[i], v + 180, f"{v:,}".replace(",", ".") if lang == "pt" else f"{v:,}",
               color=c, size=fs, bold=is_peak or is_curr)

    # Observed / projected divider
    i25   = years.index(2025)
    i26   = years.index(2026)
    divx  = (pos[i25] + pos[i26]) / 2
    ax.axvline(divx, color="#DDDDDD", lw=0.9, ls="--", zorder=2)
    obs = "Observado" if lang == "pt" else "Observed"
    prj = "Projetado" if lang == "pt" else "Projected"
    _label(ax, divx - 0.12, 14_400, obs, color=C_LIGHT, size=7.5, ha="right", bold=False)
    _label(ax, divx + 0.12, 14_400, prj, color=C_GREEN,  size=7.5, ha="left",  bold=False)

    # −30 % bracket
    i28 = years.index(2028)
    ax.hlines(5_731, pos[i25], pos[i28] + 0.45, colors="#BBBBBB", lw=1.0, ls=":", zorder=4)
    bx = pos[i28] + 0.7
    ax.annotate("", xy=(bx, 4_000), xytext=(bx, 5_731),
                arrowprops=dict(arrowstyle="<->", color=C_DARK, lw=1.0))
    ax.text(bx + 0.15, (4_000 + 5_731) / 2, "−30%",
            va="center", fontsize=9.5, color=C_DARK, fontweight="bold")

    tgt = "Meta" if lang == "pt" else "Target"
    _label(ax, pos[i28], 4_000 + 350, tgt, color=C_GREEN, size=8, bold=False)

    plt.tight_layout(pad=0.2)
    return _buf(fig)


# ── Chart 2: By biome (horizontal bars) ─────────────────────────────────────

def chart_by_biome(lang: str) -> io.BytesIO:
    data   = BIOMES_KM2[lang]
    pairs  = sorted(data.items(), key=lambda x: x[1], reverse=True)
    biomes = [p[0] for p in pairs]
    values = [p[1] for p in pairs]

    bar_colors = []
    for b in biomes:
        if "Cerrado" in b:
            bar_colors.append(C_RED)
        elif "Amazon" in b or "Amazônia" in b:
            bar_colors.append(C_BLUE)
        else:
            bar_colors.append(C_GRAY)

    _style()
    fig, ax = plt.subplots(figsize=(8.8, 3.5))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    for sp in ax.spines.values():
        sp.set_visible(False)
    ax.xaxis.set_visible(False)
    ax.yaxis.set_tick_params(length=0)
    ax.set_axisbelow(True)

    pos = np.arange(len(biomes))
    ax.barh(pos, values, height=0.6, color=bar_colors, zorder=3, linewidth=0)
    ax.set_yticks(pos)
    ax.set_yticklabels(biomes, fontsize=10.5, color=C_DARK)
    ax.invert_yaxis()
    ax.set_xlim(0, max(values) * 1.28)

    for i, (v, c) in enumerate(zip(values, bar_colors)):
        ax.text(v + max(values) * 0.015, i,
                f"{v:,}".replace(",", ".") if lang == "pt" else f"{v:,}",
                va="center", ha="left", fontsize=9.5,
                color=c if c != C_GRAY else C_MED,
                fontweight="bold" if c != C_GRAY else "normal")

    plt.tight_layout(pad=0.2)
    return _buf(fig)


# ── Chart 3: Forest cover remaining (%) ─────────────────────────────────────

def chart_forest_cover(lang: str) -> io.BytesIO:
    data   = COVER_PCT[lang]
    pairs  = sorted(data.items(), key=lambda x: x[1])
    biomes = [p[0] for p in pairs]
    values = [p[1] for p in pairs]

    bar_colors = []
    for v in values:
        if v < 40:
            bar_colors.append(C_RED)
        elif v < 65:
            bar_colors.append(C_ORANGE)
        else:
            bar_colors.append(C_GREEN)

    _style()
    fig, ax = plt.subplots(figsize=(8.8, 3.5))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    for sp in ax.spines.values():
        sp.set_visible(False)
    ax.xaxis.set_visible(False)
    ax.yaxis.set_tick_params(length=0)

    pos = np.arange(len(biomes))
    ax.barh(pos, values, height=0.6, color=bar_colors, alpha=0.88, zorder=3, linewidth=0)
    ax.set_yticks(pos)
    ax.set_yticklabels(biomes, fontsize=10.5, color=C_DARK)
    ax.set_xlim(0, 118)

    # 100 % reference
    ax.axvline(100, color="#DDDDDD", lw=0.8, ls="--", zorder=2)
    ref = "Cobertura original" if lang == "pt" else "Original cover"
    ax.text(99, len(biomes) - 0.6, ref, ha="right", va="top", fontsize=7, color=C_LIGHT)

    for i, (v, c) in enumerate(zip(values, bar_colors)):
        ax.text(v + 1.5, i, f"{v:.0f}%", va="center", ha="left",
                fontsize=9.5, color=c, fontweight="bold")

    plt.tight_layout(pad=0.2)
    return _buf(fig)


# ── Chart 4: Target trajectory ───────────────────────────────────────────────

def chart_target_trajectory(lang: str) -> io.BytesIO:
    hist_y = sorted(AMAZON_KM2)
    hist_v = [AMAZON_KM2[y] for y in hist_y]
    all_y  = hist_y + [2026, 2028]
    all_v  = hist_v + [4_866, 4_000]
    colors = [C_GRAY] * len(hist_y) + [C_GREEN, C_GREEN]

    pos = np.arange(len(all_y), dtype=float)
    fig, ax = _fig(9.2, 3.9)
    ax.bar(pos, all_v, width=0.72, color=colors, zorder=3, linewidth=0)
    ax.set_ylim(0, 15_200)
    ax.set_xticks(pos)
    ax.set_xticklabels([str(y) for y in all_y], fontsize=9, color=C_MED)

    # Dashed target line
    i25 = all_y.index(2025)
    i26 = all_y.index(2026)
    i28 = all_y.index(2028)
    tx = [pos[i25], pos[i26], pos[i28]]
    ty = [5_731,    4_866,    4_000]
    ax.plot(tx, ty, color=C_GREEN, lw=2.2, ls="--", zorder=5,
            marker="o", markersize=5, markerfacecolor=C_GREEN)
    ax.fill_between(tx, ty, 0, alpha=0.05, color=C_GREEN, zorder=1)

    for x, y in zip(tx, ty):
        _label(ax, x, y + 280, f"{y:,}".replace(",", ".") if lang == "pt" else f"{y:,}",
               color=C_GREEN, size=9, bold=True)

    tgt = "Trajetória da meta" if lang == "pt" else "Target trajectory"
    ax.text(pos[i28] + 0.12, 4_000 - 700, f"← {tgt}",
            ha="left", va="top", fontsize=8, color=C_GREEN, style="italic")

    plt.tight_layout(pad=0.2)
    return _buf(fig)


# ── Chart 5: International comparison ───────────────────────────────────────

def chart_international(lang: str) -> io.BytesIO:
    pairs   = sorted(INTL_LOSS_KM2.items(), key=lambda x: x[1], reverse=True)
    countries = [p[0] for p in pairs]
    values    = [p[1] for p in pairs]

    bar_colors = [C_BLUE if c == "Brazil" else C_GRAY for c in countries]

    _style()
    fig, ax = plt.subplots(figsize=(8.8, 3.2))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    for sp in ax.spines.values():
        sp.set_visible(False)
    ax.xaxis.set_visible(False)
    ax.yaxis.set_tick_params(length=0)

    pos = np.arange(len(countries))
    ax.barh(pos, values, height=0.55, color=bar_colors, zorder=3, linewidth=0)
    ax.set_yticks(pos)
    ax.set_yticklabels(countries, fontsize=10.5, color=C_DARK)
    ax.invert_yaxis()
    ax.set_xlim(0, max(values) * 1.3)

    for i, (v, c) in enumerate(zip(values, bar_colors)):
        ax.text(v + max(values) * 0.015, i,
                f"{v:,}".replace(",", ".") if lang == "pt" else f"{v:,}",
                va="center", ha="left", fontsize=9.5,
                color=C_BLUE if c == C_BLUE else C_MED,
                fontweight="bold" if c == C_BLUE else "normal")

    plt.tight_layout(pad=0.2)
    return _buf(fig)


# ---------------------------------------------------------------------------
# SLIDE BUILDING HELPERS
# ---------------------------------------------------------------------------

def _blank(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _rect(slide, l, t, w, h, fill_hex, line_hex=None):
    """Add a filled rectangle shape."""
    from pptx.util import Pt as _Pt
    shape = slide.shapes.add_shape(1, _in(l), _in(t), _in(w), _in(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(fill_hex)
    if line_hex:
        shape.line.color.rgb = _rgb(line_hex)
        shape.line.width = int(0.5 * 12_700)
    else:
        shape.line.fill.background()
    return shape


def _tb(slide, text: str, l, t, w, h, *,
        size=10, bold=False, italic=False, color=C_DARK,
        align=1, wrap=True, font="Arial"):
    """Add a text box. align: 1=left 2=center 3=right."""
    box = slide.shapes.add_textbox(_in(l), _in(t), _in(w), _in(h))
    tf  = box.text_frame
    tf.word_wrap = wrap
    for i, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text  = line
        run.font.size   = Pt(size)
        run.font.bold   = bold
        run.font.italic = italic
        run.font.color.rgb = _rgb(color)
        run.font.name   = font
    return box


def _pic(slide, buf: io.BytesIO, l, t, w):
    """Embed a PNG from BytesIO at given position and width."""
    slide.shapes.add_picture(buf, _in(l), _in(t), width=_in(w))


def _src(slide, text: str, color=C_LIGHT):
    _tb(slide, text, 0.35, 5.32, 9.3, 0.28, size=6.5, color=color)


def _section(slide, text: str, lang: str):
    _tb(slide, text, 0.35, 0.18, 9.3, 0.28, size=7, bold=True,
        color=_LANG_COLOR[lang])


def _divider(slide, lang: str):
    """Thin horizontal rule under the section tag."""
    _rect(slide, 0.35, 0.47, 9.3, 0.005, "#E8E8E8")


def _headline(slide, text: str, t=0.52, size=13.5):
    _tb(slide, text, 0.35, t, 9.3, 0.85, size=size, bold=True, color=C_DARK)


def _sub(slide, text: str, t=1.35):
    _tb(slide, text, 0.35, t, 9.3, 0.38, size=8.5, color=C_MED)


def _lang_bar(slide, lang: str):
    """Thin color bar at the very top and bottom."""
    lc = _LANG_COLOR[lang]
    _rect(slide, 0, 0, _SW, 0.055, lc)
    _rect(slide, 0, _SH - 0.055, _SW, 0.055, lc)


# ---------------------------------------------------------------------------
# SLIDE BUILDERS  (one function per unique slide content)
# ---------------------------------------------------------------------------

def s_cover(prs: Presentation, lang: str) -> None:
    sl = _blank(prs)
    lc = _LANG_COLOR[lang]
    _lang_bar(sl, lang)

    ll = "PT-BR" if lang == "pt" else "EN-US"
    _tb(sl, ll, 0.35, 0.12, 1.5, 0.25, size=7, bold=True, color=lc)

    _tb(sl, COPY["cover_title"][lang], 0.55, 0.7, 8.5, 1.9,
        size=30, bold=True, color=C_DARK, font="Georgia")

    _rect(sl, 0.55, 2.75, 2.8, 0.055, lc)

    _tb(sl, COPY["cover_sub"][lang],    0.55, 2.92, 8.5, 0.42, size=10, color=C_MED)
    _tb(sl, COPY["cover_credit"][lang], 0.55, 3.48, 8.5, 0.38, size=9,
        italic=True, color=C_LIGHT)


def s_lead_stat(prs: Presentation, lang: str) -> None:
    sl = _blank(prs)
    lc = _LANG_COLOR[lang]
    _lang_bar(sl, lang)
    _section(sl, COPY["sec_amazon"][lang], lang)
    _divider(sl, lang)

    _tb(sl, "56%", 0.35, 0.55, 3.8, 1.5,
        size=88, bold=True, color=C_GREEN, wrap=False)

    _headline(sl, COPY["stat_headline"][lang], t=1.9, size=15)
    _tb(sl, COPY["stat_body"][lang], 0.35, 2.68, 9.3, 1.0, size=10.5, color=C_MED)
    _tb(sl, COPY["stat_ctx"][lang],  0.35, 3.9,  9.3, 0.35, size=8.5, italic=True, color=C_LIGHT)
    _src(sl, COPY["src_prodes"][lang])


def s_amazon_historical(prs: Presentation, lang: str) -> None:
    sl  = _blank(prs)
    buf = chart_amazon_historical(lang)
    _lang_bar(sl, lang)
    _section(sl, COPY["sec_amazon"][lang], lang)
    _divider(sl, lang)
    _headline(sl, COPY["hist_headline"][lang])
    _sub(sl, COPY["hist_sub"][lang])
    _pic(sl, buf, 0.2, 1.62, 9.6)
    _src(sl, COPY["src_prodes"][lang])


def s_by_biome(prs: Presentation, lang: str) -> None:
    sl  = _blank(prs)
    buf = chart_by_biome(lang)
    _lang_bar(sl, lang)
    _section(sl, COPY["sec_biomes"][lang], lang)
    _divider(sl, lang)
    _headline(sl, COPY["biome_headline"][lang])
    _sub(sl, COPY["biome_sub"][lang])
    _pic(sl, buf, 0.4, 1.62, 9.2)
    _src(sl, COPY["src_prodes"][lang])


def s_cerrado_spotlight(prs: Presentation, lang: str) -> None:
    sl = _blank(prs)
    lc = _LANG_COLOR[lang]
    _lang_bar(sl, lang)
    _section(sl, COPY["sec_cerrado"][lang], lang)
    _divider(sl, lang)
    _headline(sl, COPY["cerrado_headline"][lang])

    # Two stat cards side by side
    for i, (num_key, lbl_key) in enumerate([
        ("cerrado_card1_num", "cerrado_card1_label"),
        ("cerrado_card2_num", "cerrado_card2_label"),
    ]):
        x = 0.4 + i * 4.75
        _rect(sl, x, 1.48, 4.35, 1.95, "#FFF8E1", C_RED)
        _tb(sl, COPY[num_key][lang], x + 0.25, 1.58, 3.85, 0.9,
            size=38, bold=True, color=C_RED)
        _tb(sl, COPY[lbl_key][lang], x + 0.25, 2.38, 3.85, 0.9,
            size=9.5, color=C_MED)

    _tb(sl, COPY["cerrado_note"][lang], 0.35, 3.55, 9.3, 1.0, size=9.5, color=C_MED)
    _src(sl, COPY["src_prodes"][lang] + "  ·  MapBiomas 2023")


def s_forest_cover(prs: Presentation, lang: str) -> None:
    sl  = _blank(prs)
    buf = chart_forest_cover(lang)
    _lang_bar(sl, lang)
    _section(sl, COPY["sec_cover"][lang], lang)
    _divider(sl, lang)
    _headline(sl, COPY["cover_headline"][lang])
    _sub(sl, COPY["cover_sub"][lang])
    _pic(sl, buf, 0.4, 1.62, 9.2)
    _src(sl, COPY["src_mapbiomas"][lang])


def s_target_2028(prs: Presentation, lang: str) -> None:
    sl  = _blank(prs)
    buf = chart_target_trajectory(lang)
    _lang_bar(sl, lang)
    _section(sl, COPY["sec_target"][lang], lang)
    _divider(sl, lang)
    _headline(sl, COPY["target_headline"][lang])
    _sub(sl, COPY["target_sub"][lang])
    _pic(sl, buf, 0.2, 1.62, 9.6)
    _src(sl, COPY["src_prodes"][lang])


def s_international(prs: Presentation, lang: str) -> None:
    sl  = _blank(prs)
    buf = chart_international(lang)
    _lang_bar(sl, lang)
    _section(sl, COPY["sec_global"][lang], lang)
    _divider(sl, lang)
    _headline(sl, COPY["intl_headline"][lang])
    _sub(sl, COPY["intl_sub"][lang])
    _pic(sl, buf, 0.4, 1.62, 9.2)
    _tb(sl, COPY["intl_note"][lang], 0.35, 4.85, 9.3, 0.35, size=6.5, italic=True, color=C_LIGHT)
    _src(sl, COPY["src_gfw"][lang])


def s_causes(prs: Presentation, lang: str) -> None:
    sl = _blank(prs)
    _lang_bar(sl, lang)
    _section(sl, COPY["sec_causes"][lang], lang)
    _divider(sl, lang)
    _headline(sl, COPY["causes_headline"][lang])

    items = COPY["causes_items"][lang]
    for i, (icon, color, title, body) in enumerate(items):
        y = 1.48 + i * 1.25
        _tb(sl, f"{icon}  {title}", 0.35, y, 9.3, 0.38,
            size=11, bold=True, color=color)
        _tb(sl, body, 0.35, y + 0.37, 9.3, 0.75, size=9.5, color=C_MED)


def s_takeaways(prs: Presentation, lang: str) -> None:
    sl = _blank(prs)
    lc = _LANG_COLOR[lang]

    # Header band
    _rect(sl, 0, 0, _SW, 1.05, lc)
    _tb(sl, COPY["sec_takeaway"][lang], 0.35, 0.08, 9.3, 0.28,
        size=7, bold=True, color="#FFFFFF")
    _tb(sl, COPY["takeaway_headline"][lang], 0.35, 0.33, 9.3, 0.62,
        size=14, bold=True, color="#FFFFFF")

    # Bottom bar
    _rect(sl, 0, _SH - 0.055, _SW, 0.055, lc)

    # Three message cards
    cards = COPY["takeaways"][lang]
    for i, (title, body) in enumerate(cards):
        y = 1.12 + i * 1.42
        _rect(sl, 0.28, y, 9.44, 1.28, "#F5F5F5", "#E0E0E0")
        # Number indicator
        _rect(sl, 0.28, y, 0.38, 1.28, lc)
        _tb(sl, str(i + 1), 0.29, y + 0.38, 0.36, 0.5,
            size=18, bold=True, color="#FFFFFF", align=2)
        _tb(sl, title, 0.75, y + 0.1,  8.85, 0.42, size=11,  bold=True, color=C_DARK)
        _tb(sl, body,  0.75, y + 0.52, 8.85, 0.65, size=9.5, color=C_MED)


# ---------------------------------------------------------------------------
# SLIDE SEQUENCE
# ---------------------------------------------------------------------------

BUILDERS = [
    s_cover,
    s_lead_stat,
    s_amazon_historical,
    s_by_biome,
    s_cerrado_spotlight,
    s_forest_cover,
    s_target_2028,
    s_international,
    s_causes,
    s_takeaways,
]

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main() -> None:
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"\n{SEP}")
    print(f"  PRODES Press Briefing Generator  v{__version__}  |  {now}")
    print(f"{SEP}\n")

    prs = Presentation()
    prs.slide_width  = _in(_SW)
    prs.slide_height = _in(_SH)

    total = len(BUILDERS) * 2
    n = 0

    for lang in ("pt", "en"):
        label = "PT-BR" if lang == "pt" else "EN-US"
        print(f"  [{label}]")
        for fn in BUILDERS:
            n += 1
            name = fn.__name__.replace("s_", "")
            print(f"    [{n:2d}/{total}]  {name}", end=" ", flush=True)
            try:
                fn(prs, lang)
                print("✓")
            except Exception as exc:
                print(f"⚠  WARN: {exc} — blank slide inserted")
                _blank(prs)

    out = Path(__file__).parent / str(CONFIG["output_path"])
    prs.save(str(out))

    print(f"\n{DIV}")
    print(f"  Saved  : {out.resolve()}")
    print(f"  Slides : {len(prs.slides)}  ({len(BUILDERS)} PT-BR + {len(BUILDERS)} EN-US)")
    print(f"{SEP}\n")


if __name__ == "__main__":
    main()
