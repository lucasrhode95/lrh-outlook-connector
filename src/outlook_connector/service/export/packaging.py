"""Output packaging (requirements v4 §10.2): exactly one download.

A flat .txt only when the result is a single text file with no attachment files. Otherwise one
.zip: text files at the root, each text file's attachments in a sibling folder named after it.
"""

from __future__ import annotations

import time
import zipfile
from dataclasses import dataclass, field
from pathlib import Path, PurePath

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


def package(files: list[TextFile], *, base_name: str) -> Package:
    if not files:
        raise ValueError("Nothing to package.")
    out_dir = exports_dir()
    stamp = time.strftime("%Y%m%d-%H%M%S")
    if len(files) == 1 and not files[0].attachments:
        target = out_dir / f"{stamp} {files[0].name}"
        target.write_text(files[0].text, encoding="utf-8")
        return Package(target, files[0].name, "text/plain; charset=utf-8", target.stat().st_size)
    filename = f"{base_name}.zip"
    target = out_dir / f"{stamp} {filename}"
    with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for item in files:
            archive.writestr(item.name, item.text.encode("utf-8"))
            for name, source in item.attachments:
                archive.write(source, f"{item.folder}/{name}")
    return Package(target, filename, "application/zip", target.stat().st_size)
