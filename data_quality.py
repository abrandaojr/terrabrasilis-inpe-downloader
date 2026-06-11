from __future__ import annotations

import hashlib
import json
import logging
import os
import tempfile
import time
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from statistics import mean, pstdev
from typing import Any, Iterable, Mapping, Sequence


@dataclass(frozen=True)
class FreshnessPolicy:
    """Freshness expectation for an input or output data asset."""

    max_age_hours: float
    severity: str = "error"


@dataclass(frozen=True)
class DataContract:
    """Schema-as-contract between producer and consumer pipeline stages."""

    name: str
    producer: str
    consumers: tuple[str, ...]
    required_columns: tuple[str, ...] = ()
    optional_columns: tuple[str, ...] = ()
    numeric_columns: tuple[str, ...] = ()
    categorical_columns: tuple[str, ...] = ()
    freshness: FreshnessPolicy | None = None
    min_rows: int | None = None
    max_null_rate: float = 0.20
    notes: str = ""


@dataclass
class LineageRecord:
    """Lightweight lineage record for one pipeline stage."""

    stage_name: str
    upstream_sources: list[str]
    transformation: str
    downstream_outputs: list[str]
    contracts: list[str] = field(default_factory=list)


@dataclass
class StageMetrics:
    """Machine-readable observability metrics for one stage."""

    stage_name: str
    status: str
    started_at: str
    finished_at: str
    duration_seconds: float
    input_row_count: int | None = None
    output_row_count: int | None = None
    null_rates: dict[str, float] = field(default_factory=dict)
    volume_anomalies: list[str] = field(default_factory=list)
    distribution_anomalies: list[str] = field(default_factory=list)
    schema_anomalies: list[str] = field(default_factory=list)
    freshness_anomalies: list[str] = field(default_factory=list)


class StageTimer:
    """Context manager used to time a named data pipeline stage."""

    def __init__(self, stage_name: str) -> None:
        self.stage_name = stage_name
        self.started_at = utc_now_iso()
        self._t0 = time.perf_counter()

    def finish(
        self,
        status: str,
        input_row_count: int | None = None,
        output_row_count: int | None = None,
        null_rates: dict[str, float] | None = None,
        anomalies: Mapping[str, list[str]] | None = None,
    ) -> StageMetrics:
        finished_at = utc_now_iso()
        anomalies = anomalies or {}
        return StageMetrics(
            stage_name=self.stage_name,
            status=status,
            started_at=self.started_at,
            finished_at=finished_at,
            duration_seconds=round(time.perf_counter() - self._t0, 3),
            input_row_count=input_row_count,
            output_row_count=output_row_count,
            null_rates=null_rates or {},
            volume_anomalies=anomalies.get("volume", []),
            distribution_anomalies=anomalies.get("distribution", []),
            schema_anomalies=anomalies.get("schema", []),
            freshness_anomalies=anomalies.get("freshness", []),
        )


class JsonLineLogger:
    """Append-only JSON logger for pipeline observability events."""

    def __init__(self, path: Path) -> None:
        self.path = path
        ensure_dir(path.parent)

    def emit(self, event: str, **fields: Any) -> None:
        payload = {"timestamp": utc_now_iso(), "event": event, **fields}
        with self.path.open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(payload, ensure_ascii=False, sort_keys=True) + "\n")


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def configure_json_logging(log_path: Path) -> JsonLineLogger:
    """Configure stdlib logging and return a JSONL event logger."""
    ensure_dir(log_path.parent)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s %(message)s",
        force=False,
    )
    return JsonLineLogger(log_path)


def require_existing_dir(path: Path, label: str) -> None:
    if not path.exists():
        raise SystemExit(f"[FATAL] {label} directory not found: {path}")
    if not path.is_dir():
        raise SystemExit(f"[FATAL] {label} path is not a directory: {path}")


def atomic_write_text(path: Path, text: str, encoding: str = "utf-8") -> None:
    ensure_dir(path.parent)
    fd, tmp_name = tempfile.mkstemp(
        prefix=f".{path.name}.", suffix=".tmp", dir=str(path.parent)
    )
    tmp = Path(tmp_name)
    try:
        with os.fdopen(fd, "w", encoding=encoding, newline="\n") as fh:
            fh.write(text)
        tmp.replace(path)
    except Exception:
        tmp.unlink(missing_ok=True)
        raise


def atomic_write_json(path: Path, payload: dict[str, Any]) -> None:
    atomic_write_text(
        path,
        json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True),
    )


def to_jsonable(value: Any) -> Any:
    if hasattr(value, "__dict__") and not isinstance(value, type):
        try:
            return asdict(value)
        except TypeError:
            pass
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, tuple):
        return list(value)
    if isinstance(value, list):
        return [to_jsonable(v) for v in value]
    if isinstance(value, dict):
        return {str(k): to_jsonable(v) for k, v in value.items()}
    return value


def file_fingerprint(path: Path, sample_bytes: int = 1_048_576) -> dict[str, Any]:
    st = path.stat()
    h = hashlib.sha256()
    with path.open("rb") as fh:
        h.update(fh.read(sample_bytes))
    return {
        "path": str(path),
        "bytes": st.st_size,
        "mtime_utc": datetime.fromtimestamp(st.st_mtime, timezone.utc).isoformat(
            timespec="seconds"
        ),
        "sha256_head": h.hexdigest(),
    }


def freshness_metrics(paths: Iterable[Path], policy: FreshnessPolicy | None) -> dict[str, Any]:
    """Compute last-updated metadata and policy violations for files."""
    now_ts = datetime.now(timezone.utc).timestamp()
    newest_ts: float | None = None
    oldest_ts: float | None = None
    checked = 0
    missing: list[str] = []
    stale: list[str] = []

    for path in paths:
        if not path.exists():
            missing.append(str(path))
            continue
        checked += 1
        mtime = path.stat().st_mtime
        newest_ts = mtime if newest_ts is None else max(newest_ts, mtime)
        oldest_ts = mtime if oldest_ts is None else min(oldest_ts, mtime)
        if policy is not None:
            age_hours = (now_ts - mtime) / 3600
            if age_hours > policy.max_age_hours:
                stale.append(str(path))

    return {
        "checked_count": checked,
        "missing": missing,
        "stale": stale,
        "newest_mtime_utc": (
            datetime.fromtimestamp(newest_ts, timezone.utc).isoformat(timespec="seconds")
            if newest_ts is not None
            else None
        ),
        "oldest_mtime_utc": (
            datetime.fromtimestamp(oldest_ts, timezone.utc).isoformat(timespec="seconds")
            if oldest_ts is not None
            else None
        ),
        "policy": to_jsonable(policy) if policy else None,
    }


def enforce_freshness(paths: Iterable[Path], policy: FreshnessPolicy, label: str) -> dict[str, Any]:
    metrics = freshness_metrics(paths, policy)
    if metrics["missing"] or (metrics["stale"] and policy.severity == "error"):
        raise SystemExit(
            f"[FATAL] Freshness check failed for {label}: "
            f"{len(metrics['missing'])} missing, {len(metrics['stale'])} stale."
        )
    return metrics


def validate_nonempty_files(paths: Iterable[Path], label: str) -> list[dict[str, Any]]:
    missing: list[str] = []
    empty: list[str] = []
    fingerprints: list[dict[str, Any]] = []

    for path in paths:
        if not path.exists():
            missing.append(str(path))
            continue
        if not path.is_file():
            missing.append(str(path))
            continue
        if path.stat().st_size <= 0:
            empty.append(str(path))
            continue
        fingerprints.append(file_fingerprint(path))

    if missing or empty:
        details = []
        if missing:
            details.append(f"missing: {len(missing)}")
        if empty:
            details.append(f"empty: {len(empty)}")
        raise SystemExit(f"[FATAL] Invalid {label} artifact(s): {', '.join(details)}")

    return fingerprints


def file_inventory(paths: Sequence[Path]) -> dict[str, Any]:
    """Return volume-style metrics for a collection of files."""
    sizes = [p.stat().st_size for p in paths if p.exists() and p.is_file()]
    return {
        "file_count": len(paths),
        "nonempty_file_count": sum(1 for size in sizes if size > 0),
        "total_bytes": sum(sizes),
        "min_bytes": min(sizes) if sizes else None,
        "max_bytes": max(sizes) if sizes else None,
        "mean_bytes": round(mean(sizes), 2) if sizes else None,
    }


def compare_volume(
    current_count: int,
    previous_count: int | None,
    label: str,
    max_drop_pct: float = 0.20,
) -> list[str]:
    """Flag unexpected count drops between runs."""
    if previous_count in (None, 0):
        return []
    drop_pct = (previous_count - current_count) / previous_count
    if drop_pct > max_drop_pct:
        return [
            f"{label} dropped {drop_pct:.1%}: previous={previous_count}, current={current_count}"
        ]
    return []


def numeric_distribution(values: Sequence[float]) -> dict[str, Any]:
    clean = [float(v) for v in values if v is not None]
    if not clean:
        return {"count": 0, "null_rate": 1.0}
    return {
        "count": len(clean),
        "min": min(clean),
        "max": max(clean),
        "mean": mean(clean),
        "stddev": pstdev(clean) if len(clean) > 1 else 0.0,
    }


def null_rate(nulls: int, total: int) -> float:
    return round(nulls / total, 6) if total else 0.0


def parquet_quality_profile(
    paths: Sequence[Path],
    contract: DataContract,
    max_files: int = 25,
) -> dict[str, Any]:
    """
    Profile Parquet files for schema, volume, null rates, and simple distributions.

    Uses pyarrow only when available because it is already part of this codebase.
    """
    try:
        import pyarrow.parquet as pq
        import pyarrow.compute as pc
    except Exception as exc:
        return {"status": "skipped", "reason": f"pyarrow unavailable: {exc}"}

    schema_anomalies: list[str] = []
    distribution_anomalies: list[str] = []
    row_count = 0
    files_profiled = 0
    column_types: dict[str, str] = {}
    null_rates: dict[str, float] = {}
    numeric_stats: dict[str, dict[str, Any]] = {}
    categorical_stats: dict[str, dict[str, Any]] = {}

    for path in paths[:max_files]:
        try:
            meta = pq.read_metadata(path)
            row_count += meta.num_rows
            files_profiled += 1
            for field in meta.schema.to_arrow_schema():
                column_types.setdefault(field.name, str(field.type))
        except Exception as exc:
            schema_anomalies.append(f"{path}: metadata read failed: {exc}")

    available = set(column_types)
    missing_required = sorted(set(contract.required_columns) - available)
    if missing_required:
        schema_anomalies.append(
            f"{contract.name}: missing required columns {missing_required}"
        )

    unexpected = sorted(
        available - set(contract.required_columns) - set(contract.optional_columns)
    )

    columns_to_read = sorted(
        (set(contract.numeric_columns) | set(contract.categorical_columns)) & available
    )
    if columns_to_read:
        try:
            table = pq.read_table([str(p) for p in paths[:max_files]], columns=columns_to_read)
            total_rows = table.num_rows
            for col in columns_to_read:
                arr = table[col]
                nr = null_rate(arr.null_count, total_rows)
                null_rates[col] = nr
                if nr > contract.max_null_rate:
                    distribution_anomalies.append(
                        f"{contract.name}.{col}: null_rate {nr:.2%} exceeds "
                        f"{contract.max_null_rate:.2%}"
                    )
                if col in contract.numeric_columns:
                    casted = pc.cast(arr, "double", safe=False)
                    numeric_stats[col] = {
                        "null_rate": nr,
                        "min": pc.min(casted).as_py(),
                        "max": pc.max(casted).as_py(),
                        "mean": pc.mean(casted).as_py(),
                        "stddev": pc.stddev(casted).as_py(),
                    }
                if col in contract.categorical_columns:
                    categorical_stats[col] = {
                        "null_rate": nr,
                        "unique_count": len(pc.unique(arr).to_pylist()),
                    }
        except Exception as exc:
            distribution_anomalies.append(f"{contract.name}: profiling failed: {exc}")

    volume_anomalies = []
    if contract.min_rows is not None and row_count < contract.min_rows:
        volume_anomalies.append(
            f"{contract.name}: row_count {row_count} below min_rows {contract.min_rows}"
        )

    return {
        "status": "ok",
        "contract": to_jsonable(contract),
        "files_profiled": files_profiled,
        "row_count": row_count,
        "column_types": column_types,
        "unexpected_columns": unexpected,
        "null_rates": null_rates,
        "numeric_stats": numeric_stats,
        "categorical_stats": categorical_stats,
        "schema_anomalies": schema_anomalies,
        "volume_anomalies": volume_anomalies,
        "distribution_anomalies": distribution_anomalies,
    }


def duplicate_names(paths: Iterable[Path]) -> dict[str, list[str]]:
    seen: dict[str, list[str]] = {}
    for path in paths:
        seen.setdefault(path.name, []).append(str(path))
    return {name: values for name, values in seen.items() if len(values) > 1}


def write_run_report(
    report_dir: Path,
    script_name: str,
    payload: dict[str, Any],
) -> Path:
    ensure_dir(report_dir)
    stem = Path(script_name).stem
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S_%f")
    path = report_dir / f"{stem}_quality_{stamp}.json"
    atomic_write_json(
        path,
        {
            "generated_at": utc_now_iso(),
            "script": script_name,
            **to_jsonable(payload),
        },
    )
    return path
