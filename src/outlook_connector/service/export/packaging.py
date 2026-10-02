"""Output packaging (requirements v4 §10.2): exactly one download.

A flat .txt only when the result is a single text file with no attachment files. Otherwise one
.zip: text files at the root, each text file's attachments in a sibling folder named after it.
"""

from __future__ import annotations

import time
import zipfile
from dataclasses import dataclass, field
from pathlib import Path, PurePath
from typing import IO, Any

from outlook_connector import config

KEEP_EXPORTS_SECONDS = 7 * 24 * 3600


@dataclass
class TextFile:
    name: str  # final file name, unique within the export
    text: str
    attachments: list[tuple[str, Path]] = field(default_factory=list)  # (name inside <stem>/, source file)

    @property
    def folder(self) -> str:
        return PurePath(self.name).stem


@dataclass
class Package:
    path: Path
    filename: str
    content_type: str
    size: int


def exports_dir() -> Path:
    directory = config.data_dir() / "exports"
    directory.mkdir(parents=True, exist_ok=True)
    cutoff = time.time() - KEEP_EXPORTS_SECONDS
    for old in directory.iterdir():
        if old.is_file() and old.stat().st_mtime < cutoff:
            old.unlink(missing_ok=True)
    return directory


def _create(out_dir: Path, filename: str, **open_args: Any) -> tuple[Path, IO[Any]]:
    """A new file that no other export (in this or another process) can also claim."""
    stamp = time.strftime("%Y%m%d-%H%M%S")
    stem, suffix = PurePath(filename).stem, PurePath(filename).suffix
    counter = 1
    while True:
        name = f"{stamp} {filename}" if counter == 1 else f"{stamp} {stem} ({counter}){suffix}"
        target = out_dir / name
        try:
            return target, target.open(**open_args)
        except FileExistsError:
            counter += 1


def package(files: list[TextFile], *, base_name: str) -> Package:
    if not files:
        raise ValueError("Nothing to package.")
    out_dir = exports_dir()
    if len(files) == 1 and not files[0].attachments:
        target, handle = _create(out_dir, files[0].name, mode="x", encoding="utf-8")
        with handle:
            handle.write(files[0].text)
        return Package(target, files[0].name, "text/plain; charset=utf-8", target.stat().st_size)
    filename = f"{base_name}.zip"
    target, handle = _create(out_dir, filename, mode="xb")
    with handle, zipfile.ZipFile(handle, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for item in files:
            archive.writestr(item.name, item.text.encode("utf-8"))
            for name, source in item.attachments:
                archive.write(source, f"{item.folder}/{name}")
    return Package(target, filename, "application/zip", target.stat().st_size)
