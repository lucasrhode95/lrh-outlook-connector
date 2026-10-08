"""Raw files for agents and the UI: attachments and MIME (.eml) saved under the data directory.

Agents receive raw files and read them with their own tools (requirements v4 §9). Files are kept
for a week, like exports.
"""

from __future__ import annotations

import tempfile
from email.parser import BytesHeaderParser
from email.policy import default
from pathlib import Path

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
        """Save one attachment as a local file.

        Entry point: the attachment is looked up on the message (unknown ids raise NotFound; cloud
        attachments are refused) before anything is downloaded.
        """
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
