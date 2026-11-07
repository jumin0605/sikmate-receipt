"""Decide which on-disk jobs can continue after a process restart.

This module does not delete files. A restart that wiped the data directory
simply has nothing to resume.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

_JOB_ID = re.compile(r"[0-9a-f]{16}")
_OPEN_STATES = {"queued", "running"}


def classify_on_restart(jobs_dir: Path) -> tuple[list[str], list[str]]:
    """Return (resumable ids, stalled ids).

    Resumable jobs still have source files and were queued or running.
    Stalled jobs were queued or running, but the originals are gone.
    Finished jobs and unreadable status files are left untouched.
    """
    if not jobs_dir.is_dir():
        return [], []
    resumable: list[str] = []
    stalled: list[str] = []
    for path in sorted(jobs_dir.iterdir()):
        if not path.is_dir() or not _JOB_ID.fullmatch(path.name):
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
        source = path / "원본"
        has_source = False
        if source.is_dir():
            try:
                has_source = any(source.iterdir())
            except OSError:
                has_source = False
        if has_source:
            resumable.append(path.name)
        else:
            stalled.append(path.name)
    return resumable, stalled
