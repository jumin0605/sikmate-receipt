"""Load project-local secrets without an extra dotenv dependency.

Explicit process environment variables take precedence. This file never logs
values and only reads configuration from the project root.
"""
from __future__ import annotations

import os
import re
from pathlib import Path
from typing import MutableMapping


def load_project_env(project_dir: Path, environ: MutableMapping[str, str] | None = None) -> None:
    target = os.environ if environ is None else environ
    env_path = project_dir / ".env"
    if not env_path.is_file():
        return
    for raw in env_path.read_text(encoding="utf-8-sig").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[7:].lstrip()
        key, separator, value = line.partition("=")
        key, value = key.strip(), value.strip()
        if not separator or not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", key):
            continue
        if len(value) >= 2 and value[0] in "\"'" and value[-1] == value[0]:
            value = value[1:-1]
        elif " #" in value:
            value = value.split(" #", 1)[0].rstrip()
        target.setdefault(key, value)
