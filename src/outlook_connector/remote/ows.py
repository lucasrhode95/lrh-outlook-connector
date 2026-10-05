"""Outlook Web's JSON RPC (OWS, ``/owa/service.svc``): the write path (architecture §5.5).

A gap fill: Graph cannot write mail with the usable first-party clients (research §2), so every
write goes here with the ``write`` token. Only this module knows OWS JSON. Contracts are the ones
proven in research §4.1–4.2; anything else is marked where it is used.

- Bearer only: no cookies, no canary. Payloads up to 2,048 URL-encoded characters travel in the
  ``X-OWA-UrlPostData`` header with an empty body, larger ones in the body.
- Every write is sent once. A missing or server-error answer raises ``WriteOutcomeUnknown``.
- Item results carry ``ResponseClass`` / ``ResponseCode``; anything but success is an error that
  names the code.
- The inbox-rule actions use a second style (research §4.4, ``Ows.call_request``): the request object
  is posted as is, and the answer reports ``WasSuccessful`` / ``ErrorCode``.
"""

from __future__ import annotations

import html
import json
import re
import uuid
from typing import Any, Protocol
from urllib.parse import quote

from outlook_connector.domain.errors import NotFound, Upstream, WriteOutcomeUnknown
from outlook_connector.domain.models import EmailProposal
from outlook_connector.remote import ids
from outlook_connector.remote.ports import FolderTarget
from outlook_connector.remote.transport import Transport, operation

OWS_URL = "https://outlook.cloud.microsoft/owa/service.svc"
SERVER_VERSION = "V2018_01_08"
URL_POST_DATA_LIMIT = 2048
MAX_ERROR_TEXT = 200
SUCCESS = frozenset({"Success", "Warning"})


class _Claims(Protocol):
    def claims(self) -> dict[str, Any]: ...


class WriteTokens(Protocol):
    def get_token(self, profile: str) -> _Claims: ...


class Ows:
    """One OWS action call: envelope, headers, item results."""

    def __init__(self, transport: Transport, tokens: WriteTokens, *, profile: str = "write") -> None:
        self._transport = transport
        self._tokens = tokens
        self._profile = profile

    def account(self) -> dict[str, Any]:
        """Claims of the write token (tid, oid, upn): the account writes act as."""
        return self._tokens.get_token(self._profile).claims()

    async def call(self, action: str, body: dict[str, Any], *, strict: bool = True) -> list[dict[str, Any]]:
        """Send one action; return its item results, in request order. ``strict``: raise unless
        every item succeeded (otherwise the caller reads each item's outcome)."""
        envelope = {
            "__type": f"{action}JsonRequest:#Exchange",
            "Header": {"__type": "JsonRequestHeaders:#Exchange", "RequestServerVersion": SERVER_VERSION},
            "Body": {"__type": f"{action}Request:#Exchange", **body},
        }
        return _items(await self._post(action, envelope), action, strict=strict)

    async def call_request(
        self, action: str, fields: dict[str, Any], *, time_zone: str | None = None
    ) -> dict[str, Any]:
        """Send one action in the bare-request style of the inbox-rule actions (research §4.4): the
        request object itself, no ``JsonRequest`` wrapper and no ``Body``. Return the answer, which
        reports ``WasSuccessful`` / ``ErrorCode`` instead of item results. Sent once, never retried.

        ``time_zone``: a Windows time zone id for ``Header.TimeZoneContext`` (Outlook Web sends one).
        """
        header: dict[str, Any] = {
            "__type": "JsonRequestHeaders:#Exchange",
            "RequestServerVersion": SERVER_VERSION,
        }
        if time_zone:
            header["TimeZoneContext"] = {
                "__type": "TimeZoneContext:#Exchange",
                "TimeZoneDefinition": {"__type": "TimeZoneDefinitionType:#Exchange", "Id": time_zone},
            }
        data = await self._post(action, {"__type": f"{action}Request:#Exchange", "Header": header, **fields})
        if not isinstance(data, dict) or not isinstance(data.get("WasSuccessful"), bool):
            raise WriteOutcomeUnknown(
                f"Outlook answered {action} without a success flag, so it is unclear whether the change "
                "was made; it was not retried. Check the mailbox before trying again."
            )
        if not data["WasSuccessful"] or data.get("ErrorCode") not in (0, None):
            text = re.sub(r"\s+", " ", str(data.get("ErrorMessage") or "")).strip()[:MAX_ERROR_TEXT]
            raise Upstream(
                f"Outlook refused {action} (error {data.get('ErrorCode')})" + (f": {text}" if text else ".")
            )
        return data

    async def _post(self, action: str, payload: dict[str, Any]) -> Any:
        """POST one payload to the action, with Outlook Web's headers. Sent once, never retried."""
        encoded = quote(json.dumps(payload, separators=(",", ":"), ensure_ascii=False), safe="-_.!~*'()")
        claims = self.account()
        upn = claims.get("upn") or claims.get("preferred_username")
        headers = {
            "Action": action,
            "X-OWA-ActionSource": action,
            "X-OWA-CorrelationId": str(uuid.uuid4()),
            "X-OWA-SessionId": str(uuid.uuid4()),
            "Prefer": 'IdType="ImmutableId"',
            "Content-Type": "application/json; charset=utf-8",
        }
        if upn:
            headers["X-AnchorMailbox"] = f"AAD-SMTP:{upn}"
        in_header = len(encoded) <= URL_POST_DATA_LIMIT
        if in_header:
            headers["X-OWA-UrlPostData"] = encoded
        return await self._transport.json(
            "POST",
            f"{OWS_URL}?action={action}&app=Mail",
            profile=self._profile,
            headers=headers,
            json_body=None if in_header else payload,
            write=True,
        )


def _items(data: Any, action: str, *, strict: bool) -> list[dict[str, Any]]:
    body = data.get("Body") if isinstance(data, dict) else None
    messages = body.get("ResponseMessages") if isinstance(body, dict) else None
    items = messages.get("Items") if isinstance(messages, dict) else None
    if not isinstance(items, list) or not items:  # the write may have happened: never a plain failure
        raise WriteOutcomeUnknown(
            f"Outlook answered {action} without item results, so it is unclear whether the change "
            "was made; it was not retried. Check the mailbox before trying again."
        )
    if not strict:
        return [item if isinstance(item, dict) else {} for item in items]
    for item in items:
        if not isinstance(item, dict) or item.get("ResponseClass") not in SUCCESS:
            raise item_error(item if isinstance(item, dict) else {}, action)
    return items


def item_error(item: dict[str, Any], action: str) -> Exception:
    code = str(item.get("ResponseCode") or "no response code")
    text = re.sub(r"\s+", " ", str(item.get("MessageText") or "")).strip()[:MAX_ERROR_TEXT]
    detail = f"{action}: {item.get('ResponseClass') or 'Error'}, {code}" + (f": {text}" if text else "")
    if code in ("ErrorItemNotFound", "ErrorInvalidIdMalformed", "ErrorInvalidIdNotAnItemAttachmentId"):
        return NotFound(f"Outlook did not find the item ({detail}).")
    return Upstream(f"Outlook refused the change ({detail}).")


def succeeded(item: dict[str, Any]) -> bool:
    return item.get("ResponseClass") in SUCCESS


def outcome(item: dict[str, Any]) -> str | None:
    """None when the item succeeded, else its response code (e.g. ErrorItemNotFound)."""
    return None if succeeded(item) else str(item.get("ResponseCode") or item.get("ResponseClass") or "Error")


def _address(address: str) -> dict[str, str]:
    return {"__type": "EmailAddress:#Exchange", "EmailAddress": address, "RoutingType": "SMTP"}


def _body(text: str) -> dict[str, str]:
    return {"__type": "BodyContentType:#Exchange", "BodyType": "Text", "Value": text}


def _html_body(text: str) -> dict[str, str]:
    """Plain text as HTML that shows exactly the same text: escaped, line breaks kept, and repeated
    spaces, tabs and leading spaces kept as non-breaking spaces (HTML would collapse them; Outlook's
    desktop renderer ignores CSS ``white-space``). A reply needs an HTML body: Exchange then quotes
    the original's HTML untouched, with its inline images (a Text body flattens it)."""
    lines = html.escape(text).replace("\r\n", "\n").split("\n")
    return {
        "__type": "BodyContentType:#Exchange",
        "BodyType": "HTML",
        "Value": "<div>" + "<br>".join(map(_keep_spaces, lines)) + "</div>",
    }


def _keep_spaces(line: str) -> str:
    line = line.replace("\t", "&nbsp;" * 4)
    line = re.sub(r" {2,}", lambda run: "&nbsp;" * (len(run.group()) - 1) + " ", line)
    return "&nbsp;" + line[1:] if line.startswith(" ") else line


def _created_id(items: list[dict[str, Any]]) -> str | None:
    inner = items[0].get("Items")
    first = inner[0] if isinstance(inner, list) and inner and isinstance(inner[0], dict) else {}
    item_id = first.get("ItemId")
    value = item_id.get("Id") if isinstance(item_id, dict) else None
    return ids.to_graph(value) if isinstance(value, str) and value else None


class OwsMailWriter:
    """Implements ``MailWriter`` over OWS.

    Assumes (not re-checked here): proposals come from ``Writes.propose`` (addresses, subject and body
    validated), message ids are Graph immutable ids from this connector, folder targets were resolved by
    the service, and the write account was checked (``Writes.check_account``). Nothing is re-validated
    here; every call is sent once.
    """

    def __init__(self, ows: Ows) -> None:
        self._ows = ows

    def account(self) -> dict[str, Any]:
        return self._ows.account()

    async def create_draft(self, message: EmailProposal) -> str | None:
        """Save into Drafts (never sends). Returns the draft's Graph id when Outlook reports it."""
        with operation("saving a draft"):
            return _created_id(await self._ows.call("CreateItem", _create(message, "SaveOnly")))

    async def send(self, message: EmailProposal) -> None:
        """Send once and keep a copy in Sent Items. Never retried."""
        with operation("sending a message"):
            await self._ows.call("CreateItem", _create(message, "SendAndSaveCopy"))

    async def send_draft(self, draft_id: str, subject: str) -> None:
        """Send an existing draft as it is, the way Outlook Web does: ``UpdateItem`` with
        ``SendAndSaveCopy`` (``SendItem`` is not supported over OWS; live 2026-10-04). The update
        sets the subject the draft already has. Sent once, never retried."""
        body = {
            "ItemChanges": [
                {
                    "__type": "ItemChange:#Exchange",
                    "ItemId": _item_id(draft_id),
                    "Updates": [
                        {
                            "__type": "SetItemField:#Exchange",
                            "Path": {"__type": "PropertyUri:#Exchange", "FieldURI": "item:Subject"},
                            "Item": {"__type": "Message:#Exchange", "Subject": subject},
                        }
                    ],
                }
            ],
            "ConflictResolution": "AlwaysOverwrite",
            "MessageDisposition": "SendAndSaveCopy",
            "SavedItemFolderId": {
                "__type": "TargetFolderId:#Exchange",
                "BaseFolderId": {"__type": "DistinguishedFolderId:#Exchange", "Id": "sentitems"},
            },
            "SuppressReadReceipts": True,
            "SendCalendarInvitationsOrCancellations": "SendToNone",
        }
        with operation("sending a reply"):
            await self._ows.call("UpdateItem", body)

    # ---------------------------------------------------------------- mutations (research §4.2)
    # Each returns {message id: None when done, else Outlook's response code}, one request per call
    # (the service keeps calls small). Ids are Graph immutable ids; they survive moves.

    async def set_read(self, message_ids: list[str], is_read: bool) -> dict[str, str | None]:
        with operation("changing read state"):
            changes = {mid: ("message:IsRead", {"IsRead": is_read}) for mid in message_ids}
            return await self._update(changes)

    async def set_flag(self, message_ids: list[str], flagged: bool) -> dict[str, str | None]:
        status = "Flagged" if flagged else "NotFlagged"
        flag = {"Flag": {"__type": "FlagType:#Exchange", "FlagStatus": status}}
        with operation("changing flags"):
            return await self._update({mid: ("item:Flag", flag) for mid in message_ids})

    async def move(self, message_ids: list[str], folder: FolderTarget) -> dict[str, str | None]:
        body = {
            "ToFolderId": {"__type": "TargetFolderId:#Exchange", "BaseFolderId": _folder(folder)},
            "ItemIds": [_item_id(mid) for mid in message_ids],
            "ReturnNewItemIds": True,
        }
        with operation("moving messages"):
            return _per_id(message_ids, await self._ows.call("MoveItem", body, strict=False))

    async def delete(self, message_ids: list[str]) -> dict[str, str | None]:
        """Move to Deleted Items (``MoveToDeletedItems``). There is no hard delete."""
        body = {
            "ItemIds": [_item_id(mid) for mid in message_ids],
            "DeleteType": "MoveToDeletedItems",
            "SendMeetingCancellations": "SendToNone",
            "AffectedTaskOccurrences": "AllOccurrences",
            "SuppressReadReceipts": True,
        }
        with operation("deleting messages"):
            return _per_id(message_ids, await self._ows.call("DeleteItem", body, strict=False))

    async def _update(self, changes: dict[str, tuple[str, dict[str, Any]]]) -> dict[str, str | None]:
        """``UpdateItem`` with one ``SetItemField`` per message, as proven (research §4.2)."""
        body = {
            "ItemChanges": [
                {
                    "__type": "ItemChange:#Exchange",
                    "ItemId": _item_id(mid),
                    "Updates": [
                        {
                            "__type": "SetItemField:#Exchange",
                            "Path": {"__type": "PropertyUri:#Exchange", "FieldURI": field_uri},
                            "Item": {"__type": "Message:#Exchange", **props},
                        }
                    ],
                }
                for mid, (field_uri, props) in changes.items()
            ],
            "ConflictResolution": "AlwaysOverwrite",
            "MessageDisposition": "SaveOnly",
            "SuppressReadReceipts": True,
            "SendCalendarInvitationsOrCancellations": "SendToNone",
        }
        return _per_id(list(changes), await self._ows.call("UpdateItem", body, strict=False))


def _folder(target: FolderTarget) -> dict[str, str]:
    if target.well_known:
        return {"__type": "DistinguishedFolderId:#Exchange", "Id": target.well_known}
    return {"__type": "FolderId:#Exchange", "Id": ids.to_ows(target.folder_id)}


def _item_id(message_id: str) -> dict[str, str]:
    return {"__type": "ItemId:#Exchange", "Id": ids.to_ows(message_id)}


def _per_id(message_ids: list[str], items: list[dict[str, Any]]) -> dict[str, str | None]:
    if len(items) != len(message_ids):  # the change may have been made: read back, never repeat
        raise WriteOutcomeUnknown(
            f"Outlook answered {len(items)} item results for {len(message_ids)} messages; it was not retried."
        )
    return {mid: outcome(item) for mid, item in zip(message_ids, items, strict=True)}


def _create(message: EmailProposal, disposition: str) -> dict[str, Any]:
    """``CreateItem`` as proven by the 2026-10-01 self-send (research §4.2), with the disposition
    switched: ``SendAndSaveCopy`` sends, ``SaveOnly`` leaves a draft in Drafts.

    Replies use EWS's ``ReplyToItem`` / ``ReplyAllToItem`` response objects, which append the
    quoted original; recipients and subject are passed explicitly, so the result matches the
    proposal. A reply's own text goes in an HTML body, so the quoted original keeps its formatting
    and inline images (live 2026-10-04). New messages stay plain text (W7).

    Assumes (not re-checked here): ``message`` comes from ``Writes.propose``.
    """
    recipients = {
        "ToRecipients": [_address(a) for a in message.to],
        "CcRecipients": [_address(a) for a in message.cc],
        "BccRecipients": [_address(a) for a in message.bcc],
    }
    if message.reply_to_message_id:
        kind = "ReplyAllToItem" if message.reply_all else "ReplyToItem"
        item: dict[str, Any] = {
            "__type": f"{kind}:#Exchange",
            "ReferenceItemId": {"__type": "ItemId:#Exchange", "Id": ids.to_ows(message.reply_to_message_id)},
            "NewBodyContent": _html_body(message.body),
            "Subject": message.subject,
            **recipients,
        }
        compose = "replyAll" if message.reply_all else "reply"
    else:
        item = {
            "__type": "Message:#Exchange",
            "Body": _body(message.body),
            "From": {},
            **recipients,
            "Subject": message.subject,
            "Importance": "Normal",
            "IsDeliveryReceiptRequested": False,
            "IsReadReceiptRequested": False,
            "IsSendIndividually": False,
            "MessageDisposition": disposition,
            "ShouldIgnoreChangeKey": True,
            "operation": "New",
        }
        compose = "newMail"
    return {
        "ClientSupportsIrm": True,
        "ComposeOperation": compose,
        "MessageDisposition": disposition,
        "Items": [item],
        "TimeFormat": "",
        "SendOnNotFoundError": True,
        "RemoteExecute": True,
        "ShouldSuppressReadReceipt": True,
        "OutboundCharset": "AutoDetect",
        "ItemShape": {
            "__type": "ItemResponseShape:#Exchange",
            "BaseShape": "IdOnly",
            "ClientSupportsIrm": True,
            "AdditionalProperties": [{"__type": "PropertyUri:#Exchange", "FieldURI": "ItemLastModifiedTime"}],
        },
        "ShapeName": "MailCompose",
    }
