"""
Centralized logging configuration.

Terminal format : 2026-02-19 16:18:00 | INFO     | module_name | message
File hierarchy  : logs/<year>/<month>/<date>/<job_id>.log
"""
import logging
import os
from datetime import datetime


# ---------------------------------------------------------------------------
# Formatter
# ---------------------------------------------------------------------------

class PipeFormatter(logging.Formatter):
    """
    Produces pipe-delimited log lines with fixed-width level column.

    Example:
        2026-02-19 16:18:00 | INFO     | queue_manager | Processing job …
    """

    LEVEL_WIDTH = 8  # 'CRITICAL' is 8 chars — the longest standard level

    def format(self, record: logging.LogRecord) -> str:
        ts = datetime.fromtimestamp(record.created).strftime("%Y-%m-%d %H:%M:%S")
        level = record.levelname.ljust(self.LEVEL_WIDTH)
        # Strip the 'job_' prefix inserted by per-job loggers so the module
        # column remains readable (e.g. "job_job_1_abc" → "job_1_abc")
        module = record.name
        return f"{ts} | {level} | {module} | {record.getMessage()}"


# ---------------------------------------------------------------------------
# Root-level bootstrap
# ---------------------------------------------------------------------------

def configure_root_logger(level: int = logging.INFO) -> None:
    """
    Configure the root logger with a PipeFormatter console handler.
    Call this once at application startup (in main.py / lifespan).
    """
    root = logging.getLogger()
    root.setLevel(level)

    # Avoid duplicate handlers if called more than once
    if any(isinstance(h, logging.StreamHandler) for h in root.handlers):
        return

    handler = logging.StreamHandler()
    handler.setLevel(level)
    handler.setFormatter(PipeFormatter())
    root.addHandler(handler)


# ---------------------------------------------------------------------------
# Per-job file logger
# ---------------------------------------------------------------------------

def get_job_file_handler(job_id: str, base_log_dir: str = "logs") -> logging.FileHandler:
    """
    Return a FileHandler that writes to:
        <base_log_dir>/<year>/<month>/<date>/<job_id>.log

    The directory is created automatically.
    """
    now = datetime.now()
    log_dir = os.path.join(
        base_log_dir,
        str(now.year),
        f"{now.month:02d}",
        f"{now.day:02d}",
    )
    os.makedirs(log_dir, exist_ok=True)

    log_path = os.path.join(log_dir, f"{job_id}.log")
    handler = logging.FileHandler(log_path, mode="w", encoding="utf-8")
    handler.setLevel(logging.DEBUG)
    handler.setFormatter(PipeFormatter())
    return handler
