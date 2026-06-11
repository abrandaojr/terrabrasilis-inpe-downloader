from __future__ import annotations

import os
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]


def _env_path(name: str, default: Path) -> Path:
    value = os.environ.get(name)
    if value:
        return Path(value).expanduser().resolve()
    return default.resolve()


WORKSPACE_ROOT = _env_path("PRODES_HOME", PROJECT_ROOT / "workspace")
ZIP_ROOT = _env_path("PRODES_ZIP_ROOT", WORKSPACE_ROOT / "zip")
GEOPARQUET_DIR = _env_path("PRODES_GEOPARQUET_DIR", WORKSPACE_ROOT / "geoparquet")
TABLES_DIR = _env_path("PRODES_TABLES_DIR", WORKSPACE_ROOT / "tables")
FIGURES_DIR = _env_path("PRODES_FIGURES_DIR", WORKSPACE_ROOT / "figures")
REPORTS_DIR = _env_path("PRODES_REPORTS_DIR", WORKSPACE_ROOT / "reports")
PRESENTATIONS_DIR = _env_path("PRODES_PRESENTATIONS_DIR", WORKSPACE_ROOT / "presentations")
EXTRACT_DIR = _env_path("PRODES_EXTRACT_DIR", WORKSPACE_ROOT / "extracted")


PIPELINE_DIRS = (
    WORKSPACE_ROOT,
    ZIP_ROOT,
    GEOPARQUET_DIR,
    TABLES_DIR,
    FIGURES_DIR,
    REPORTS_DIR,
    PRESENTATIONS_DIR,
)


def ensure_pipeline_dirs(include_extract: bool = False) -> None:
    dirs = PIPELINE_DIRS + ((EXTRACT_DIR,) if include_extract else ())
    for path in dirs:
        path.mkdir(parents=True, exist_ok=True)
