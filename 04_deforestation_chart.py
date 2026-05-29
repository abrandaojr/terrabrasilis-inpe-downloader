import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np

data = {
    2015: 6207, 2016: 7893, 2017: 6947, 2018: 7536,
    2019: 10129, 2020: 10851, 2021: 13038,
    2022: 11594, 2023: 9064, 2024: 6518,
    2025: 5731,
}
targets = {2026: 4866, 2028: 4000}

GREEN = "#2E7D32"
LGRAY = "#CCCCCC"
DGRAY = "#444444"

years  = sorted(data) + [2026, 2028]
values = [data[y] for y in sorted(data)] + [4866, 4000]
colors = [LGRAY] * len(data) + [GREEN, GREEN]
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

# x axis
ax.set_xticks(pos)
ax.set_xticklabels([str(y) for y in years], fontsize=10, color="#888888")
ax.xaxis.set_tick_params(length=0)

# y axis — hidden, values on bars do the job
ax.set_ylim(0, 16200)
ax.yaxis.set_visible(False)

# value labels on ALL bars
important = {2021, 2025, 2028}
for i, (yr, val) in enumerate(zip(years, values)):
    is_important = yr in important
    is_green     = yr in (2026, 2028)
    txt_color    = GREEN if is_green else DGRAY
    fw           = "bold" if is_important else "normal"
    fs           = 10.5  if is_important else 9.0
    fmt          = f"{val:,}"
    ax.text(pos[i], val + 220, fmt,
            ha="center", va="bottom",
            fontsize=fs, color=txt_color, fontweight=fw)

# baseline label
i25 = years.index(2025)
ax.text(pos[i25], data[2025] + 220 + 900, "Baseline",
        ha="center", va="bottom", fontsize=8.5, color=DGRAY, style="italic")

# target label
i28 = years.index(2028)
ax.text(pos[i28], targets[2028] + 220 + 900, "Target",
        ha="center", va="bottom", fontsize=8.5, color=GREEN, style="italic")

# reference line 2025 → 2028
ax.hlines(5731, pos[i25], pos[i28] + 0.32,
          colors="#BBBBBB", lw=1.0, ls=":", zorder=4)

# -30% bracket
bx = pos[i28] + 0.52
ax.annotate("", xy=(bx, 4000), xytext=(bx, 5731),
            arrowprops=dict(arrowstyle="<->", color=DGRAY, lw=1.1))
ax.text(bx + 0.18, (4000 + 5731) / 2, "−30%",
        va="center", fontsize=10.5, color=DGRAY, fontweight="bold")

# observed / projected divider — after 2025, before 2026
sep = (pos[i25] + pos[i25 + 1]) / 2
ax.axvline(sep, color="#DDDDDD", lw=0.9, ls="--", zorder=2)
ax.text(sep - 0.15, 15600, "Observed",
        ha="right", va="top", fontsize=9, color="#BBBBBB", style="italic")
ax.text(sep + 0.15, 15600, "Projected",
        ha="left",  va="top", fontsize=9, color=GREEN, style="italic")

# title
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
plt.savefig("amazon_deforestation_norad.png",
            dpi=220, bbox_inches="tight", facecolor="white")
print("saved")