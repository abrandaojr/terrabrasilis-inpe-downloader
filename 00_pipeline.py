"""
00_pipeline.py
==============
Orchestrates the full PRODES data pipeline:

  Step 1 — 01_download_zips.py          Download ZIPs from TerraBrasilis
  Step 2 — 02_convert_to_geoparquet.py  Convert vectors → GeoParquet, rasters → COG
  Step 3 — 03_deforestation_chart.py    Generate deforestation rate chart
  Step 4 — 04_generate_presentation.py  Generate bilingual press PowerPoint

Usage
-----
    python 00_pipeline.py              # run all steps, stop on first failure
    python 00_pipeline.py -k           # run all steps, continue past failures
    python 00_pipeline.py --from 2     # resume from step 2
    python 00_pipeline.py --steps 1 3  # run only steps 1 and 3
    python 00_pipeline.py --steps 2 3 4 -k  # run steps 2-4, keep going on error

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
    (4, "04_generate_presentation.py", "Generate bilingual press PowerPoint"),
]

SEP  = "=" * 65
DIV  = "-" * 65
HERE = Path(__file__).parent

# ---------------------------------------------------------------------------
# Argument parsing (no external dependency)
# ---------------------------------------------------------------------------

def _parse_args() -> tuple[set[int], bool]:
    """Return (step numbers to run, keep_going flag)."""
    args      = sys.argv[1:]
    all_steps = {s[0] for s in STEPS}
    keep_going = "-k" in args
    args = [a for a in args if a != "-k"]

    if not args:
        return all_steps, keep_going

    if args[0] == "--from" and len(args) >= 2:
        start = int(args[1])
        return {n for n in all_steps if n >= start}, keep_going

    if args[0] == "--steps":
        return {int(n) for n in args[1:]}, keep_going

    print(__doc__)
    sys.exit(0)

# ---------------------------------------------------------------------------
# Step runner
# ---------------------------------------------------------------------------

def _run_step(
    n: int,
    script: str,
    label: str,
    step_idx: int,
    total: int,
    keep_going: bool,
) -> tuple[float, bool]:
    """Run one step. Returns (elapsed_seconds, success)."""
    print(f"\n{SEP}")
    print(f"  STEP {n}/{len(STEPS)}  [{step_idx}/{total} selected]  —  {label}")
    print(f"{SEP}\n")

    t0 = time.perf_counter()
    try:
        result = subprocess.run([sys.executable, str(HERE / script)])
    except KeyboardInterrupt:
        raise
    elapsed = time.perf_counter() - t0

    if result.returncode != 0:
        print(f"\n{SEP}")
        print(f"  [FAILED]  Step {n} exited with code {result.returncode}.")
        if keep_going:
            print(f"  Continuing to next step  (-k flag is set).")
        else:
            print(f"  Fix the issue and resume with:  python 00_pipeline.py --from {n}")
            print(f"  Or run all remaining steps:     python 00_pipeline.py --from {n} -k")
        print(SEP)
        if not keep_going:
            sys.exit(result.returncode)
        return elapsed, False

    print(f"\n  Step {n} done  ({elapsed:.1f}s)")
    return elapsed, True

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    selected, keep_going = _parse_args()
    to_run   = [s for s in STEPS if s[0] in selected]
    skipped  = [s for s in STEPS if s[0] not in selected]
    total    = len(to_run)
    pipeline_t = time.perf_counter()

    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"\n{SEP}")
    print(
        f"  PRODES Pipeline  v{__version__}  |  {now}  |  "
        f"{total} step(s) selected" +
        ("  |  -k (keep going)" if keep_going else "")
    )
    if skipped:
        print(f"  Skipping: {', '.join(f'step {s[0]}' for s in skipped)}")
    print(SEP)

    results: dict[int, tuple[float, bool]] = {}   # n → (elapsed, success)
    for idx, (n, script, label) in enumerate(to_run, 1):
        results[n] = _run_step(n, script, label, idx, total, keep_going)

    total_elapsed = time.perf_counter() - pipeline_t
    any_failed    = any(not ok for _, ok in results.values())

    print(f"\n{SEP}")
    print(f"  {'PIPELINE COMPLETE' if not any_failed else 'PIPELINE DONE WITH ERRORS'}")
    print(DIV)
    for n, script, label in STEPS:
        if n in results:
            elapsed, ok = results[n]
            status = f"{elapsed:>6.1f}s" if ok else f"{elapsed:>6.1f}s  ← FAILED"
            print(f"  [{n}] {label:<50}  {status}")
        else:
            print(f"  [{n}] {label:<50}  skipped")
    print(DIV)
    print(f"  Total elapsed: {total_elapsed:.1f}s")
    if any_failed:
        failed_steps = [n for n, (_, ok) in results.items() if not ok]
        print(f"  Failed steps : {', '.join(str(n) for n in failed_steps)}")
        print(f"  Retry with   : python 00_pipeline.py --steps {' '.join(str(n) for n in failed_steps)}")
    print(SEP + "\n")

    if any_failed:
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n{SEP}")
        print("  Interrupted.")
        print(SEP + "\n")
        sys.exit(130)
