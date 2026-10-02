"""Output packaging (requirements v4 §10.2): exactly one download.

A flat file (.txt or .jsonl) only when the result is a single text file with no attachment files.
Otherwise one .zip: text files at the root, each text file's attachments in a sibling folder
named after it.
"""

from __future__ import annotations

import time
import zipfile
from dataclasses import dataclass, field
from pathlib import Path, PurePath

from outlook_connector.service.localfiles import claim, kept_dir

CONTENT_TYPES = {".txt": "text/plain; charset=utf-8", ".jsonl": "application/x-ndjson; charset=utf-8"}


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


def package(files: list[TextFile], *, base_name: str) -> Package:
    if not files:
        raise ValueError("Nothing to package.")
    out_dir = kept_dir("exports")
    stamp = time.strftime("%Y%m%d-%H%M%S")
    if len(files) == 1 and not files[0].attachments:
        only = files[0]
        target = claim(out_dir, f"{stamp} {only.name}")
        target.write_text(only.text, encoding="utf-8")
        content_type = CONTENT_TYPES.get(PurePath(only.name).suffix, "text/plain; charset=utf-8")
        return Package(target, only.name, content_type, target.stat().st_size)
    filename = f"{base_name}.zip"
    target = claim(out_dir, f"{stamp} {filename}")
    with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for item in files:
            archive.writestr(item.name, item.text.encode("utf-8"))
            for name, source in item.attachments:
                archive.write(source, f"{item.folder}/{name}")
    return Package(target, filename, "application/zip", target.stat().st_size)
