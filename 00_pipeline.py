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
    (1, "01_download_zips.py", "Download ZIPs from TerraBrasilis"),
    (
        2,
        "02_convert_to_geoparquet.py",
        "Convert vectors → GeoParquet | rasters → COG",
    ),
    (3, "03_deforestation_chart.py", "Generate deforestation rate chart"),
    (
        4,
        "04_generate_presentation.py",
        "Generate bilingual press PowerPoint (+maps)",
    ),
    (
        5,
        "05_organize_geoparquet.py",
        "Organize GeoParquet folder → catalog + _organized/",
    ),
    (6, "06_export_tables.py", "Export all tables to Excel → tables/"),
]

# Constants for formatting output and paths
SEP = "=" * 65
DIV = "-" * 65
HERE = Path(__file__).parent


# ---------------------------------------------------------------------------
# Argument parsing (no external dependency)
# ---------------------------------------------------------------------------


def _print_usage_and_exit(error_message: str | None = None) -> None:
    """Prints the script usage and exits with an error code."""
    if error_message:
        print(f"Error: {error_message}", file=sys.stderr)
    print(__doc__)
    sys.exit(1)


def _parse_args() -> tuple[set[int], bool]:
    """Parse command-line arguments. Returns (step numbers to run, keep_going flag)."""
    args = sys.argv[1:]
    all_step_numbers = {s[0] for s in STEPS}
    keep_going = "-k" in args
    args = [arg for arg in args if arg != "-k"]

    if not args:
        return all_step_numbers, keep_going

    if args[0] == "--from":
        if len(args) != 2:
            _print_usage_and_exit(
                "Incorrect usage of '--from'. "
                "Expected: --from <step_number>"
            )
        try:
            start_step = int(args[1])
        except ValueError:
            _print_usage_and_exit(
                f"Invalid step number for '--from': '{args[1]}'. "
                "Must be an integer."
            )
        if start_step not in all_step_numbers:
            _print_usage_and_exit(
                f"Step {start_step} does not exist. "
                f"Available steps are: {', '.join(map(str, sorted(all_step_numbers)))}."
            )
        return {n for n in all_step_numbers if n >= start_step}, keep_going

    if args[0] == "--steps":
        if len(args) < 2:
            _print_usage_and_exit(
                "Incorrect usage of '--steps'. "
                "Expected: --steps <step_number_1> [step_number_2]..."
            )
        selected_steps = set()
        for arg_step in args[1:]:
            try:
                step_num = int(arg_step)
            except ValueError:
                _print_usage_and_exit(
                    f"Invalid step number for '--steps': '{arg_step}'. "
                    "Must be an integer."
                )
            if step_num not in all_step_numbers:
                _print_usage_and_exit(
                    f"Step {step_num} does not exist. "
                    f"Available steps are: {', '.join(map(str, sorted(all_step_numbers)))}."
                )
            selected_steps.add(step_num)
        return selected_steps, keep_going

    _print_usage_and_exit(f"Unknown argument: '{args[0]}'.")


# ---------------------------------------------------------------------------
# Step runner
# ---------------------------------------------------------------------------


def _run_step(
    step_number: int,
    script_name: str,
    step_label: str,
    current_run_idx: int,
    total_steps_to_run: int,
    keep_going: bool,
) -> tuple[float, bool]:
    """
    Run one pipeline step.

    Returns:
        (elapsed_seconds, success_status)
    """
    print(f"\n{SEP}")
    print(
        f"  STEP {step_number}/{len(STEPS)}  "
        f"[{current_run_idx}/{total_steps_to_run} selected]  "
        f"—  {step_label}"
    )
    print(f"{SEP}\n")

    start_time = time.perf_counter()
    try:
        # Use sys.executable to ensure the correct Python interpreter is used
        result = subprocess.run(
            [sys.executable, str(HERE / script_name)],
            check=False,  # Don't raise CalledProcessError, check returncode
        )
    except KeyboardInterrupt:
        # Re-raise to be caught by the main handler,
        # ensuring consistent exit behavior for Ctrl+C
        raise
    step_elapsed_time = time.perf_counter() - start_time

    if result.returncode != 0:
        print(f"\n{SEP}")
        print(
            f"  [FAILED]  Step {step_number} exited with code "
            f"{result.returncode}."
        )
        if keep_going:
            print("  Continuing to next step (-k flag is set).")
        else:
            print(
                f"  Fix the issue and resume with:  "
                f"python 00_pipeline.py --from {step_number}"
            )
            print(
                f"  Or run all remaining steps:     "
                f"python 00_pipeline.py --from {step_number} -k"
            )
        print(SEP)
        if not keep_going:
            sys.exit(result.returncode)
        return step_elapsed_time, False

    print(f"\n  Step {step_number} done  ({step_elapsed_time:.1f}s)")
    return step_elapsed_time, True


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> None:
    """
    Main function to orchestrate the PRODES data pipeline.
    """
    selected_step_numbers, keep_going = _parse_args()
    steps_to_run = [s for s in STEPS if s[0] in selected_step_numbers]
    skipped_steps = [s for s in STEPS if s[0] not in selected_step_numbers]
    total_steps_to_run = len(steps_to_run)
    pipeline_start_time = time.perf_counter()

    current_time = datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"\n{SEP}")
    header_message = (
        f"  PRODES Pipeline  v{__version__}  |  {current_time}  |  "
        f"{total_steps_to_run} step(s) selected"
    )
    if keep_going:
        header_message += "  |  -k (keep going)"
    print(header_message)

    if skipped_steps:
        skipped_info = ", ".join(
            f"step {s[0]}" for s in sorted(skipped_steps, key=lambda x: x[0])
        )
        print(f"  Skipping: {skipped_info}")
    print(SEP)

    results: dict[int, tuple[float, bool]] = {}  # step_number → (elapsed, success)
    for idx, (step_num, script, label) in enumerate(steps_to_run, 1):
        results[step_num] = _run_step(
            step_num, script, label, idx, total_steps_to_run, keep_going
        )

    total_pipeline_elapsed = time.perf_counter() - pipeline_start_time
    any_failed = any(not ok for _, ok in results.values())

    print(f"\n{SEP}")
    print(f"  {'PIPELINE COMPLETE' if not any_failed else 'PIPELINE DONE WITH ERRORS'}")
    print(DIV)
    for step_num, _, label in STEPS:
        if step_num in results:
            elapsed, ok = results[step_num]
            status = (
                f"{elapsed:>6.1f}s"
                if ok
                else f"{elapsed:>6.1f}s  ← FAILED"
            )
            print(f"  [{step_num}] {label:<50}  {status}")
        else:
            print(f"  [{step_num}] {label:<50}  skipped")
    print(DIV)
    print(f"  Total elapsed: {total_pipeline_elapsed:.1f}s")
    if any_failed:
        failed_steps = [
            n for n, (_, ok) in results.items() if not ok
        ]
        failed_steps_str = " ".join(str(n) for n in sorted(failed_steps))
        print(f"  Failed steps : {failed_steps_str}")
        print(
            f"  Retry with   : python 00_pipeline.py --steps "
            f"{failed_steps_str}"
        )
    print(SEP + "\n")

    if any_failed:
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n{SEP}")
        print("  Pipeline interrupted by user (Ctrl+C).")
        print(SEP + "\n")
        sys.exit(130)  # Standard exit code for KeyboardInterrupt