from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

from prodes_config import ensure_pipeline_dirs

ROOT = Path(__file__).resolve().parent
VENV_DIR = ROOT / ".venv"


def _bin(name: str) -> Path:
    if os.name == "nt":
        return VENV_DIR / "Scripts" / name
    return VENV_DIR / "bin" / name


def _run(cmd: list[str]) -> None:
    print("+ " + " ".join(str(c) for c in cmd))
    subprocess.check_call(cmd)


def main() -> None:
    if sys.version_info < (3, 11):
        raise SystemExit("Python 3.11 or newer is required.")

    ensure_pipeline_dirs(include_extract=True)

    if not VENV_DIR.exists():
        _run([sys.executable, "-m", "venv", str(VENV_DIR)])

    python = _bin("python.exe" if os.name == "nt" else "python")
    _run([str(python), "-m", "pip", "install", "--upgrade", "pip"])
    _run([str(python), "-m", "pip", "install", "--upgrade", "uv"])
    _run([str(python), "-m", "uv", "pip", "install", "-r", str(ROOT / "requirements.txt")])

    print("\nEnvironment ready.")
    if os.name == "nt":
        print(r"Activate with: .venv\Scripts\activate")
    else:
        print("Activate with: source .venv/bin/activate")
    print("Run with: python 00_pipeline.py")


if __name__ == "__main__":
    main()
