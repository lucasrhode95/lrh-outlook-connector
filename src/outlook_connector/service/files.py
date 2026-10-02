"""Raw files for agents: attachments and MIME (.eml) saved under the data directory.

Agents receive raw files and read them with their own tools (requirements v4 §9). Files are kept
for a week, like exports.
"""

from __future__ import annotations

import time
from pathlib import Path

from pydantic import BaseModel

from outlook_connector import config
from outlook_connector.domain.errors import InvalidRequest, NotFound
from outlook_connector.service.export.attachments import dedupe, safe_name
from outlook_connector.service.mailbox import Mailbox

KEEP_SECONDS = 7 * 24 * 3600


class SavedFile(BaseModel):
    path: str
    name: str
    content_type: str | None
    size: int


def downloads_dir() -> Path:
    directory = config.data_dir() / "downloads"
    directory.mkdir(parents=True, exist_ok=True)
    cutoff = time.time() - KEEP_SECONDS
    for old in directory.iterdir():
        if old.is_file() and old.stat().st_mtime < cutoff:
            old.unlink(missing_ok=True)
    return directory


def _target(name: str) -> Path:
    directory = downloads_dir()
    taken = {p.name.lower() for p in directory.iterdir()}
    return directory / dedupe(name, taken)


class Files:
    def __init__(self, mailbox: Mailbox) -> None:
        self.mailbox = mailbox

    async def download_attachment(self, message_id: str, attachment_id: str) -> SavedFile:
        attachments = await self.mailbox.reader.list_attachments(message_id)
        attachment = next((a for a in attachments if a.id == attachment_id), None)
        if attachment is None:
            raise NotFound("No such attachment on this message.")
        if attachment.kind == "reference":
            raise InvalidRequest(
                "This is a cloud (reference) attachment; it has no file content to download."
            )
        name = safe_name(attachment.name, fallback="attachment", eml=attachment.kind == "item")
        target = _target(name)
        size = await self.mailbox.reader.download_attachment(message_id, attachment_id, target)
        content_type = "message/rfc822" if attachment.kind == "item" else attachment.content_type
        return SavedFile(path=str(target), name=target.name, content_type=content_type, size=size)

    async def save_mime(self, message_id: str) -> SavedFile:
        message = await self.mailbox.message(message_id)
        if message.is_deleted:
            raise NotFound("The message was deleted on the server; its MIME source is no longer available.")
        target = _target(safe_name(message.subject, fallback="message") + ".eml")
        size = await self.mailbox.reader.download_mime(message_id, target)
        return SavedFile(path=str(target), name=target.name, content_type="message/rfc822", size=size)
