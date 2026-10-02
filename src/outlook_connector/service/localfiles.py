"""Local output files (exports, downloads) under the data directory.

Each kind has its own folder; files older than a week are removed when the folder is next used.
A new file name is claimed with an exclusive create, so concurrent processes never share a file.
"""

from __future__ import annotations

import time
from pathlib import Path, PurePath

from outlook_connector import config

KEEP_SECONDS = 7 * 24 * 3600


def kept_dir(name: str) -> Path:
    directory = config.data_dir() / name
    directory.mkdir(parents=True, exist_ok=True)
    cutoff = time.time() - KEEP_SECONDS
    for old in directory.iterdir():
        if old.is_file() and old.stat().st_mtime < cutoff:
            old.unlink(missing_ok=True)
    return directory


def claim(directory: Path, filename: str) -> Path:
    """Create a new empty file named ``filename`` (or ``stem (2).ext``, ...) and return its path."""
    stem, suffix = PurePath(filename).stem, PurePath(filename).suffix
    counter = 1
    while True:
        target = directory / (filename if counter == 1 else f"{stem} ({counter}){suffix}")
        try:
            target.open("xb").close()
            return target
        except FileExistsError:
            counter += 1
