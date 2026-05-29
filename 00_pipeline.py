"""
00_pipeline.py
==============
Orchestrates the full PRODES data pipeline:

  Step 1 — 01_download_zips.py          Download ZIPs from TerraBrasilis
  Step 2 — 02_convert_to_geoparquet.py  Convert vectors → GeoParquet, rasters → COG
  Step 3 — 03_deforestation_chart.py    Generate deforestation rate chart

Usage
-----
    python 00_pipeline.py              # run all steps
    python 00_pipeline.py --from 2     # resume from step 2
    python 00_pipeline.py --steps 1 3  # run only steps 1 and 3

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

import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Pipeline definition
# ---------------------------------------------------------------------------

STEPS: list[tuple[int, str, str]] = [
    (1, "01_download_zips.py",         "Download ZIPs from TerraBrasilis"),
    (2, "02_convert_to_geoparquet.py", "Convert vectors → GeoParquet  |  rasters → COG"),
    (3, "03_deforestation_chart.py",   "Generate deforestation rate chart"),
]

SEP  = "=" * 65
DIV  = "-" * 65
HERE = Path(__file__).parent

# ---------------------------------------------------------------------------
# Argument parsing (no external dependency)
# ---------------------------------------------------------------------------

def _parse_args() -> set[int]:
    """Return the set of step numbers to run."""
    args = sys.argv[1:]
    all_steps = {s[0] for s in STEPS}

    if not args:
        return all_steps

    if args[0] == "--from" and len(args) >= 2:
        start = int(args[1])
        return {n for n in all_steps if n >= start}

    if args[0] == "--steps":
        return {int(n) for n in args[1:]}

    print(__doc__)
    sys.exit(0)

# ---------------------------------------------------------------------------
# Step runner
# ---------------------------------------------------------------------------

def _run_step(n: int, script: str, label: str, step_idx: int, total: int) -> float:
    print(f"\n{SEP}")
    print(f"  STEP {n}/{len(STEPS)}  [{step_idx}/{total} selected]  —  {label}")
    print(f"{SEP}\n")

    t0     = time.perf_counter()
    result = subprocess.run([sys.executable, str(HERE / script)])
    elapsed = time.perf_counter() - t0

    if result.returncode != 0:
        print(f"\n{SEP}")
        print(f"  [FAILED]  Step {n} exited with code {result.returncode}.")
        print(f"  Fix the issue and resume with:  python 00_pipeline.py --from {n}")
        print(SEP)
        sys.exit(result.returncode)

    print(f"\n  Step {n} done  ({elapsed:.1f}s)")
    return elapsed

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    selected   = _parse_args()
    to_run     = [s for s in STEPS if s[0] in selected]
    total      = len(to_run)
    skipped    = [s for s in STEPS if s[0] not in selected]
    pipeline_t = time.perf_counter()

    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"\n{SEP}")
    print(f"  PRODES Pipeline  v{__version__}  |  {now}  |  {total} step(s) selected")
    if skipped:
        print(f"  Skipping: {', '.join(f'step {s[0]}' for s in skipped)}")
    print(SEP)

    elapsed_per_step: dict[int, float] = {}
    for idx, (n, script, label) in enumerate(to_run, 1):
        elapsed_per_step[n] = _run_step(n, script, label, idx, total)

    total_elapsed = time.perf_counter() - pipeline_t

    print(f"\n{SEP}")
    print(f"  PIPELINE COMPLETE")
    print(DIV)
    for n, script, label in STEPS:
        if n in elapsed_per_step:
            print(f"  [{n}] {label:<50}  {elapsed_per_step[n]:>6.1f}s")
        else:
            print(f"  [{n}] {label:<50}  skipped")
    print(DIV)
    print(f"  Total elapsed: {total_elapsed:.1f}s")
    print(SEP + "\n")


if __name__ == "__main__":
    main()
