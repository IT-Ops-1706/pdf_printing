"""
Daily job-state logger.

Maintains a ``states.json`` file inside each day's log folder::

    logs/<year>/<month>/<day>/states.json

The file records every job that passes through the queue, including
timestamps for each status transition, duration, and error details.

This module is **observational only** — it never influences queue
behaviour or application control-flow.
"""
import json
import logging
import os
import tempfile
import threading
from datetime import datetime
from typing import Optional

logger = logging.getLogger(__name__)

_BASE_LOG_DIR = "logs"
_lock = threading.Lock()


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _today_dir() -> str:
    """Return ``logs/<year>/<month>/<day>`` for today and ensure it exists."""
    now = datetime.now()
    path = os.path.join(
        _BASE_LOG_DIR,
        str(now.year),
        f"{now.month:02d}",
        f"{now.day:02d}",
    )
    os.makedirs(path, exist_ok=True)
    return path


def _states_path() -> str:
    return os.path.join(_today_dir(), "states.json")


def _read() -> dict:
    """Read existing states.json or return a fresh skeleton."""
    path = _states_path()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as fh:
                return json.load(fh)
        except (json.JSONDecodeError, OSError) as exc:
            logger.warning("Corrupt states.json, resetting: %s", exc)
    return _skeleton()


def _skeleton() -> dict:
    return {
        "date": datetime.now().strftime("%Y-%m-%d"),
        "summary": {
            "total": 0,
            "queued": 0,
            "processing": 0,
            "completed": 0,
            "failed": 0,
        },
        "last_updated": None,
        "jobs": {},
    }


def _recalc_summary(data: dict) -> None:
    """Derive summary counters from the jobs dict (single source of truth)."""
    counts = {"queued": 0, "processing": 0, "completed": 0, "failed": 0}
    for job in data["jobs"].values():
        status = job.get("status", "queued")
        if status in counts:
            counts[status] += 1
    data["summary"] = {
        "total": len(data["jobs"]),
        **counts,
    }


def _write(data: dict) -> None:
    """Atomic write: temp-file in the same directory, then os.replace()."""
    data["last_updated"] = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    target = _states_path()
    dir_name = os.path.dirname(target)
    fd, tmp = tempfile.mkstemp(dir=dir_name, suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=2, ensure_ascii=False)
        os.replace(tmp, target)
    except Exception:
        # Best-effort cleanup on failure
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


def _safe(func):
    """Decorator: swallow exceptions so logging never disrupts the pipeline."""
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as exc:
            logger.error("job_state_logger.%s failed: %s", func.__name__, exc)
    return wrapper


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

@_safe
def record_queued(job_id: str, filename: str) -> None:
    """Record a newly queued job."""
    with _lock:
        data = _read()
        data["jobs"][job_id] = {
            "filename": filename,
            "status": "queued",
            "queued_at": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
            "processing_at": None,
            "completed_at": None,
            "duration_seconds": None,
            "error": None,
            "error_type": None,
        }
        _recalc_summary(data)
        _write(data)


@_safe
def record_processing(job_id: str) -> None:
    """Mark a job as processing."""
    with _lock:
        data = _read()
        job = data["jobs"].get(job_id)
        if job is None:
            return
        job["status"] = "processing"
        job["processing_at"] = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
        _recalc_summary(data)
        _write(data)


@_safe
def record_completed(job_id: str) -> None:
    """Mark a job as completed and compute duration."""
    with _lock:
        data = _read()
        job = data["jobs"].get(job_id)
        if job is None:
            return
        now = datetime.now()
        job["status"] = "completed"
        job["completed_at"] = now.strftime("%Y-%m-%dT%H:%M:%S")
        # Duration from queued_at to completed_at
        if job.get("queued_at"):
            try:
                start = datetime.strptime(job["queued_at"], "%Y-%m-%dT%H:%M:%S")
                job["duration_seconds"] = round((now - start).total_seconds(), 2)
            except ValueError:
                pass
        _recalc_summary(data)
        _write(data)


@_safe
def record_failed(
    job_id: str,
    error: str,
    error_type: Optional[str] = None,
) -> None:
    """Mark a job as failed with error details and compute duration."""
    with _lock:
        data = _read()
        job = data["jobs"].get(job_id)
        if job is None:
            return
        now = datetime.now()
        job["status"] = "failed"
        job["completed_at"] = now.strftime("%Y-%m-%dT%H:%M:%S")
        job["error"] = error
        job["error_type"] = error_type
        if job.get("queued_at"):
            try:
                start = datetime.strptime(job["queued_at"], "%Y-%m-%dT%H:%M:%S")
                job["duration_seconds"] = round((now - start).total_seconds(), 2)
            except ValueError:
                pass
        _recalc_summary(data)
        _write(data)
