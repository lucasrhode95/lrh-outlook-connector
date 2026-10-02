"""Raw files for agents and the UI: attachments and MIME (.eml) saved under the data directory.

Agents receive raw files and read them with their own tools (requirements v4 §9). Files are kept
for a week, like exports.
"""

from __future__ import annotations

from pydantic import BaseModel

from outlook_connector.domain.errors import InvalidRequest, NotFound
from outlook_connector.service.export.attachments import safe_name
from outlook_connector.service.localfiles import claim, kept_dir
from outlook_connector.service.mailbox import Mailbox


class SavedFile(BaseModel):
    path: str
    name: str
    content_type: str | None
    size: int


class Files:
    def __init__(self, mailbox: Mailbox) -> None:
        self.mailbox = mailbox

    async def download_attachment(self, message_id: str, attachment_id: str) -> SavedFile:
        attachments = await self.mailbox.attachments(message_id)
        attachment = next((a for a in attachments if a.id == attachment_id), None)
        if attachment is None:
            raise NotFound("No such attachment on this message.")
        if attachment.kind == "reference":
            raise InvalidRequest(
                "This is a cloud (reference) attachment; it has no file content to download."
            )
        target = claim(
            kept_dir("downloads"),
            safe_name(attachment.name, fallback="attachment", eml=attachment.kind == "item"),
        )
        try:
            size = await self.mailbox.reader.download_attachment(message_id, attachment_id, target)
        except Exception:
            target.unlink(missing_ok=True)
            raise
        content_type = "message/rfc822" if attachment.kind == "item" else attachment.content_type
        return SavedFile(path=str(target), name=target.name, content_type=content_type, size=size)

    async def save_mime(self, message_id: str) -> SavedFile:
        known = self.mailbox.store.summaries([message_id]).get(message_id)
        subject = known.subject if known else (await self.mailbox.message(message_id)).subject
        target = claim(kept_dir("downloads"), safe_name(subject, fallback="message") + ".eml")
        try:
            size = await self.mailbox.reader.download_mime(message_id, target)
        except NotFound:
            target.unlink(missing_ok=True)
            if known:
                self.mailbox.store.mark_deleted([message_id])
            raise NotFound(
                "The message was deleted on the server; its MIME source is no longer available."
            ) from None
        except Exception:
            target.unlink(missing_ok=True)
            raise
        return SavedFile(path=str(target), name=target.name, content_type="message/rfc822", size=size)
