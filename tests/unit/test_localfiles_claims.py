from __future__ import annotations

import os
import time
from pathlib import Path

import pytest

from outlook_connector.service import localfiles

DAY = 24 * 3600


def _age(path: Path, days: float) -> None:
    stamp = time.time() - days * DAY
    os.utime(path, (stamp, stamp))


def test_expired_files_are_removed_and_fresh_files_are_kept(tmp_path: Path) -> None:
    folder = tmp_path / "downloads"
    folder.mkdir()
    expired = folder / "expired.bin"
    fresh = folder / "fresh.bin"
    expired.write_bytes(b"old")
    fresh.write_bytes(b"new")
    _age(expired, 8)

    localfiles.kept_dir("downloads", root=tmp_path)

    assert not expired.exists()
    assert fresh.read_bytes() == b"new"


def test_an_expired_file_held_open_elsewhere_is_kept_and_cleanup_goes_on(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    folder = tmp_path / "downloads"
    folder.mkdir()
    locked = folder / "locked.bin"
    other = folder / "other.bin"
    locked.write_bytes(b"in use")
    other.write_bytes(b"stale")
    _age(locked, 8)
    _age(other, 8)
    real_replace = os.replace

    def refuse_locked(src: str | os.PathLike[str], dst: str | os.PathLike[str]) -> None:
        if Path(src) == locked:
            raise PermissionError("in use by another process")
        real_replace(src, dst)

    monkeypatch.setattr(localfiles.os, "replace", refuse_locked)

    localfiles.kept_dir("downloads", root=tmp_path)

    assert locked.read_bytes() == b"in use"
    assert not other.exists()


def test_a_file_that_replaces_an_expired_one_is_restored_not_deleted(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    folder = tmp_path / "downloads"
    folder.mkdir()
    target = folder / "report.pdf"
    target.write_bytes(b"expired")
    _age(target, 8)
    real_replace = os.replace
    swapped = False

    def swap_before_claim(src: str | os.PathLike[str], dst: str | os.PathLike[str]) -> None:
        nonlocal swapped
        if Path(src) == target and not swapped:
            swapped = True
            target.unlink()
            target.write_bytes(b"fresh download")  # another process takes the name meanwhile
        real_replace(src, dst)

    monkeypatch.setattr(localfiles.os, "replace", swap_before_claim)

    localfiles.kept_dir("downloads", root=tmp_path)

    assert target.read_bytes() == b"fresh download"
    assert sorted(path.name for path in folder.iterdir()) == ["report.pdf"]
