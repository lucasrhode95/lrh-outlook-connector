"""Local output files (exports, downloads) under the data directory.

Each kind has its own folder. Files older than a week are removed when the folder is next used, including
files in its subfolders (the download cache keeps one subfolder per source). A new file name is claimed with
an exclusive create, so concurrent processes never share a file. One process cleans a folder at a time (a
lock file beside the folder); a process that finds the lock held skips its cleanup, since another process
is doing it. An expired file is renamed to a private name before it is deleted, so the file deleted is the
one whose age was checked. A file held open elsewhere is kept.
"""

from __future__ import annotations

import contextlib
import os
import sys
import time
import uuid
from collections.abc import Iterator
from pathlib import Path, PurePath

from outlook_connector import config

KEEP_SECONDS = 7 * 24 * 3600


def kept_dir(name: str) -> Path:
    """Return the ``name`` folder under the data directory, removing expired files."""
    directory = config.data_dir() / name
    directory.mkdir(parents=True, exist_ok=True)
    with _cleanup_lock(directory) as held:
        if held:
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


@contextlib.contextmanager
def _cleanup_lock(directory: Path) -> Iterator[bool]:
    """Hold the cleanup lock of ``directory`` and yield True, or yield False when another process holds
    it. While cleanups never overlap, an expired file can change only by its own process touching it, so
    a fresh file moved aside is put back under a name nobody else frees. The operating system releases
    the lock when its process ends, so a crash leaves no stale lock."""
    lock_path = directory.with_name(f".{directory.name}.cleanup.lock")
    with lock_path.open("a+b") as lock:
        try:
            _lock(lock.fileno())
        except OSError:  # held by another process
            yield False
            return
        try:
            yield True
        finally:
            _unlock(lock.fileno())


if sys.platform == "win32":
    import msvcrt

    def _lock(fd: int) -> None:
        os.lseek(fd, 0, os.SEEK_SET)
        msvcrt.locking(fd, msvcrt.LK_NBLCK, 1)

    def _unlock(fd: int) -> None:
        os.lseek(fd, 0, os.SEEK_SET)
        msvcrt.locking(fd, msvcrt.LK_UNLCK, 1)

else:
    import fcntl

    def _lock(fd: int) -> None:
        fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)

    def _unlock(fd: int) -> None:
        fcntl.flock(fd, fcntl.LOCK_UN)


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
