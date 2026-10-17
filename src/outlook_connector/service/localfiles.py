"""Local output files (exports, downloads) under the data directory.

Each kind has its own folder. Files older than a week are removed when the folder is next used, including
files in its subfolders (the download cache keeps one subfolder per source). A new file name is claimed with
an exclusive create, so concurrent processes never share a file. An expired file is renamed to a private
name before it is deleted, so the file deleted is the one whose age was checked. A file held open elsewhere
is kept.
"""

from __future__ import annotations

import contextlib
import os
import time
import uuid
from pathlib import Path, PurePath

from outlook_connector import config

KEEP_SECONDS = 7 * 24 * 3600


def kept_dir(name: str) -> Path:
    """Return the ``name`` folder under the data directory, removing expired files."""
    directory = config.data_dir() / name
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
            elif old.is_file():
                _remove_if_expired(old, cutoff)
        # Removed by another process, or open elsewhere (Windows): kept for a later cleanup.
        # Any other error is a real problem and still propagates.
        except (FileNotFoundError, PermissionError):
            continue


def _remove_if_expired(path: Path, cutoff: float) -> None:
    """Delete ``path`` if it has expired. It is renamed to a private name first, so a file that replaces
    it after the first check is not deleted: its age is checked again, and a fresh file goes back under
    its name."""
    if path.stat().st_mtime >= cutoff:
        return
    claimed = path.with_name(f".{path.name}.{uuid.uuid4().hex}.expiring")
    os.replace(path, claimed)
    if claimed.stat().st_mtime < cutoff:
        claimed.unlink(missing_ok=True)
        return
    with contextlib.suppress(OSError):  # the name was taken again: the file stays under its private name
        os.link(claimed, path)
        claimed.unlink()


def _remove_empty(folder: Path) -> None:
    with contextlib.suppress(OSError):  # another process may have put a file in it meanwhile
        folder.rmdir()
