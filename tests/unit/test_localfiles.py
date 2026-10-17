from __future__ import annotations

import errno
import os
from pathlib import Path

import pytest

from outlook_connector import config
from outlook_connector.service.localfiles import kept_dir


def test_kept_dir_tolerates_file_disappearing_before_stat(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(config, "data_dir", lambda: tmp_path)
    directory = tmp_path / "exports"
    directory.mkdir()
    old_file = directory / "old.txt"
    old_file.write_text("temporary", encoding="utf-8")
    real_stat = Path.stat
    calls = 0

    def disappear_before_stat(path: Path, *args: object, **kwargs: object) -> object:
        nonlocal calls
        if path == old_file:
            calls += 1
            if calls == 1:
                raise FileNotFoundError(path)
        return real_stat(path, *args, **kwargs)  # type: ignore[arg-type]

    monkeypatch.setattr(Path, "stat", disappear_before_stat)

    assert kept_dir("exports") == directory
    assert calls == 1


def test_kept_dir_does_not_suppress_other_cleanup_errors(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(config, "data_dir", lambda: tmp_path)
    directory = tmp_path / "exports"
    directory.mkdir()
    old_file = directory / "old.txt"
    old_file.write_text("temporary", encoding="utf-8")
    os.utime(old_file, (0, 0))
    real_replace = os.replace

    def failing_replace(src: str | os.PathLike[str], dst: str | os.PathLike[str]) -> None:
        if Path(src) == old_file:
            raise OSError(errno.EIO, "disk error", str(old_file))
        real_replace(src, dst)

    monkeypatch.setattr(os, "replace", failing_replace)

    with pytest.raises(OSError) as error:
        kept_dir("exports")
    assert error.value.errno == errno.EIO
