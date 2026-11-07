"""Restart decisions for one uvicorn process.

The web container runs a single process with an in-memory queue. That queue
cannot claim a job exclusively if more than one process is started. This
module does not delete files in 원본 and does not add a cross-process lock.
"""
from __future__ import annotations

import json
import re
import shutil
from pathlib import Path

_JOB_ID = re.compile(r"[0-9a-f]{16}")
_OPEN_STATES = {"queued", "running"}
ALLOWED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff", ".pdf"}
_DERIVED_NAMES = ("정리결과", "json", "지출현황.xlsx", "정리된_영수증.zip")


def _expected_count(status: dict) -> int | None:
    count = status.get("file_count")
    if isinstance(count, bool) or not isinstance(count, int) or count < 1:
        return None
    return count


def _readable_original(path: Path, allowed_ext: set[str]) -> bool:
    if path.is_symlink() or not path.is_file():
        return False
    if path.suffix.lower() not in allowed_ext:
        return False
    try:
        with path.open("rb") as handle:
            return bool(handle.read(1))
    except OSError:
        return False


def _source_problem(source: Path, expected: int | None, allowed_ext: set[str]) -> str | None:
    """Return a stall code, or None when the originals are complete."""
    if not source.is_dir() or source.is_symlink():
        return "missing"
    try:
        children = list(source.iterdir())
    except OSError:
        return "incomplete"
    if expected is None or any(child.is_symlink() or not child.is_file() for child in children):
        return "incomplete"
    readable = [child for child in children if _readable_original(child, allowed_ext)]
    if len(readable) != len(children) or len(readable) != expected:
        return "incomplete"
    return None


def classify_on_restart(
    jobs_dir: Path,
    allowed_ext: set[str] | None = None,
) -> tuple[list[str], list[tuple[str, str]]]:
    """Return (resumable ids, stalled (id, code) pairs).

    A job is resumable only when 원본 contains exactly file_count readable
    receipt files. Directories, links, partial sets, and unreadable files are
    incomplete and are not processed. Nothing is deleted.
    """
    allowed = ALLOWED_EXTENSIONS if allowed_ext is None else set(allowed_ext)
    if not jobs_dir.is_dir():
        return [], []
    resumable: list[str] = []
    stalled: list[tuple[str, str]] = []
    for path in sorted(jobs_dir.iterdir()):
        if not path.is_dir() or path.is_symlink() or not _JOB_ID.fullmatch(path.name):
            continue
        status_path = path / "status.json"
        if not status_path.is_file():
            continue
        try:
            status = json.loads(status_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError, UnicodeError):
            continue
        if not isinstance(status, dict) or status.get("state") not in _OPEN_STATES:
            continue
        problem = _source_problem(path / "원본", _expected_count(status), allowed)
        if problem:
            stalled.append((path.name, problem))
        else:
            resumable.append(path.name)
    return resumable, stalled


def clear_derived_outputs(job_dir: Path) -> list[str]:
    """Remove generated exports so a replay cannot mix them with a new run.

    원본 and status.json are left byte-for-byte. Symlinks are unlinked and
    their targets are not removed.
    """
    removed: list[str] = []
    if not job_dir.is_dir() or job_dir.is_symlink():
        return removed
    for name in _DERIVED_NAMES:
        path = job_dir / name
        if path.is_symlink():
            path.unlink()
            removed.append(name)
            continue
        if not path.exists():
            continue
        if path.is_dir():
            shutil.rmtree(path)
        elif path.is_file():
            path.unlink()
        else:
            continue
        removed.append(name)
    return removed
