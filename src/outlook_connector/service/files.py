"""Raw files for agents and the UI: attachments and MIME (.eml) saved under the data directory.

Agents receive raw files and read them with their own tools (requirements §10). Attachment metadata,
markers and downloads share file names. Files are kept for a week, like exports.
"""

from __future__ import annotations

import hashlib
import os
import tempfile
import uuid
from email.parser import BytesHeaderParser
from email.policy import default
from pathlib import Path

from pydantic import BaseModel

from outlook_connector.domain.errors import ConnectorError, InvalidRequest, NotFound
from outlook_connector.domain.models import Attachment
from outlook_connector.service.export.attachments import safe_name
from outlook_connector.service.localfiles import claim, kept_dir
from outlook_connector.service.mailbox import Mailbox
from outlook_connector.service.markers import DOWNLOADABLE_KINDS

MAX_NAMES = 100
_NOT_A_FILE = "a cloud link with no file here; open it from the message in Outlook"


class SavedFile(BaseModel):
    path: str
    name: str
    content_type: str | None
    size: int


class UnavailableFile(BaseModel):
    name: str
    reason: str


class SavedFiles(BaseModel):
    files: list[SavedFile]
    unavailable: list[UnavailableFile] = []


class Files:
    def __init__(self, mailbox: Mailbox) -> None:
        self.mailbox = mailbox

    async def download_attachment(self, message_id: str, names: list[str] | None = None) -> SavedFiles:
        """Save files by their listed names; all downloadable files when ``names`` is None.

        Entry point: ``names`` is checked against the message's files before anything is downloaded
        (case-insensitive; an unknown name raises NotFound and lists the names there are). A cloud link, or
        a file that cannot be fetched, is listed under ``unavailable`` while the others are saved; a call
        that saves nothing raises. Each file is fetched once into the download cache, keyed on the
        immutable message and attachment ids, and reused while it is kept.
        """
        attachments = await self.mailbox.attachments(message_id)
        by_name = {a.file_name.casefold(): a for a in attachments if a.file_name}
        if names is None:
            chosen = [a for a in attachments if a.file_name and a.kind in DOWNLOADABLE_KINDS]
            if not chosen:
                raise NotFound("This message has no files to download.")
        else:
            wanted = list(dict.fromkeys(name.casefold() for name in names))
            if not wanted or len(wanted) > MAX_NAMES:
                raise InvalidRequest(f"names must hold 1 to {MAX_NAMES} file names, or be omitted for all.")
            unknown = [name for name in names if name.casefold() not in by_name]
            if unknown:
                known = ", ".join(a.file_name for a in attachments if a.file_name) or "none"
                raise NotFound(f"No file named {', '.join(unknown)} on this message. Its files: {known}.")
            chosen = [by_name[name] for name in wanted]
        saved: list[SavedFile] = []
        unavailable: list[UnavailableFile] = []
        errors: list[ConnectorError] = []
        directory = kept_dir("downloads")
        for attachment in chosen:
            name = attachment.file_name or "attachment"
            if attachment.kind not in DOWNLOADABLE_KINDS:
                unavailable.append(UnavailableFile(name=name, reason=_NOT_A_FILE))
                continue
            try:
                saved.append(await self._cached(directory, message_id, attachment, name))
            except ConnectorError as exc:
                errors.append(exc)
                unavailable.append(UnavailableFile(name=name, reason=str(exc)))
        if not saved:
            if len(errors) == 1 and len(unavailable) == 1:
                raise errors[0]
            raise InvalidRequest(
                "Nothing was saved: " + "; ".join(f"{u.name}: {u.reason}" for u in unavailable)
            )
        return SavedFiles(files=saved, unavailable=unavailable)

    async def _cached(self, directory: Path, message_id: str, attachment: Attachment, name: str) -> SavedFile:
        """The attachment in the download cache, fetched only when it is not there. A file is published
        under its name only once complete, so a file found there is whole."""
        source = hashlib.sha256(f"{message_id}\n{attachment.id}".encode()).hexdigest()[:32]
        folder = directory / source
        target = folder / name
        content_type = "message/rfc822" if attachment.kind == "item" else attachment.content_type
        if target.is_file():
            os.utime(target)  # kept another week from this use
            return SavedFile(
                path=str(target), name=name, content_type=content_type, size=target.stat().st_size
            )
        folder.mkdir(parents=True, exist_ok=True)
        try:
            partial = claim(folder, f".{uuid.uuid4().hex}.part")
        except FileNotFoundError:
            # Another process's cleanup removed the folder (empty and a week old) after mkdir. A recreated
            # folder is fresh, so cleanup leaves it alone for a week: one retry is enough.
            folder.mkdir(parents=True, exist_ok=True)
            partial = claim(folder, f".{uuid.uuid4().hex}.part")
        try:
            size = await self.mailbox.reader.download_attachment(message_id, attachment.id, partial)
        except BaseException:
            partial.unlink(missing_ok=True)
            raise
        try:
            os.replace(partial, target)
        except PermissionError:  # Windows: another call published the same file and holds it open
            partial.unlink(missing_ok=True)
            if not target.is_file():
                raise
        return SavedFile(path=str(target), name=name, content_type=content_type, size=size)

    async def save_mime(self, message_id: str) -> SavedFile:
        """Save message MIME as a local .eml named from its Subject header.

        Entry point: the message id is taken as given; a message that is not on the server raises
        NotFound. The MIME download supplies both the file and its subject, so no summary lookup is
        needed.
        """
        directory = kept_dir("downloads")
        with tempfile.TemporaryDirectory(prefix=".mime-", dir=directory) as temp:
            source = Path(temp) / "message.eml"
            try:
                size = await self.mailbox.reader.download_mime(message_id, source)
            except NotFound:
                raise NotFound(
                    "The message MIME source was not found; the message may have been deleted, "
                    "moved out of reach, or the id may be wrong."
                ) from None
            header = bytearray()
            with source.open("rb") as stream:
                for line in stream:
                    header.extend(line)
                    if line in (b"\r\n", b"\n"):
                        break
            subject = BytesHeaderParser(policy=default).parsebytes(bytes(header)).get("Subject")
            target = claim(
                directory,
                safe_name(str(subject) if subject else None, fallback="message") + ".eml",
            )
            try:
                source.replace(target)
            except Exception:
                target.unlink(missing_ok=True)
                raise
        return SavedFile(path=str(target), name=target.name, content_type="message/rfc822", size=size)
