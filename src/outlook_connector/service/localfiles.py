"""Local output files (exports, downloads) under the data directory.

Each kind has its own folder. Files older than a week are removed when the folder is next used, including
files in its subfolders (the download cache keeps one subfolder per source). A new file name is claimed with
an exclusive create, so concurrent processes never share a file.
"""

from __future__ import annotations

import contextlib
import time
from pathlib import Path, PurePath

from outlook_connector import config

KEEP_SECONDS = 7 * 24 * 3600


def kept_dir(name: str, *, root: Path | None = None) -> Path:
    """Return the ``name`` folder under ``root`` (default: the data directory), removing expired files."""
    directory = (root if root is not None else config.data_dir()) / name
    directory.mkdir(parents=True, exist_ok=True)
    _remove_expired(directory, time.time() - KEEP_SECONDS)
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


def _remove_expired(directory: Path, cutoff: float) -> None:
    for old in directory.iterdir():
        try:
            if old.is_dir() and not old.is_symlink():
                _remove_expired(old, cutoff)
                if old.stat().st_mtime < cutoff and not any(old.iterdir()):
                    _remove_empty(old)
            elif old.is_file() and old.stat().st_mtime < cutoff:
                old.unlink(missing_ok=True)
        except FileNotFoundError:
            continue  # another process removed the entry after it was listed


def _remove_empty(folder: Path) -> None:
    with contextlib.suppress(OSError):  # another process may have put a file in it meanwhile
        folder.rmdir()
