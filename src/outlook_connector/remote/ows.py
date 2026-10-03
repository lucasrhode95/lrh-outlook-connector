"""Outlook Web's JSON RPC (OWS, ``/owa/service.svc``): the write path (architecture §5.5).

A gap fill: Graph cannot write mail with the usable first-party clients (research §2), so every
write goes here with the ``write`` token. Only this module knows OWS JSON. Contracts are the ones
proven in research §4.1–4.2; anything else is marked where it is used.

- Bearer only: no cookies, no canary. Payloads up to 2,048 URL-encoded characters travel in the
  ``X-OWA-UrlPostData`` header with an empty body, larger ones in the body.
- Every write is sent once. A missing or server-error answer raises ``WriteOutcomeUnknown``.
- Item results carry ``ResponseClass`` / ``ResponseCode``; anything but success is an error that
  names the code.
"""

from __future__ import annotations

import json
import re
import uuid
from typing import Any, Protocol
from urllib.parse import quote

from outlook_connector.domain.errors import NotFound, Upstream, WriteOutcomeUnknown
from outlook_connector.domain.models import EmailProposal
from outlook_connector.remote import ids
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

    async def call(self, action: str, body: dict[str, Any]) -> list[dict[str, Any]]:
        """Send one action; return its item results (all successful) or raise."""
        envelope = {
            "__type": f"{action}JsonRequest:#Exchange",
            "Header": {"__type": "JsonRequestHeaders:#Exchange", "RequestServerVersion": SERVER_VERSION},
            "Body": {"__type": f"{action}Request:#Exchange", **body},
        }
        encoded = quote(json.dumps(envelope, separators=(",", ":"), ensure_ascii=False), safe="-_.!~*'()")
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
        data = await self._transport.json(
            "POST",
            f"{OWS_URL}?action={action}&app=Mail",
            profile=self._profile,
            headers=headers,
            json_body=None if in_header else envelope,
            write=True,
        )
        return _items(data, action)


def _items(data: Any, action: str) -> list[dict[str, Any]]:
    body = data.get("Body") if isinstance(data, dict) else None
    messages = body.get("ResponseMessages") if isinstance(body, dict) else None
    items = messages.get("Items") if isinstance(messages, dict) else None
    if not isinstance(items, list) or not items:  # the write may have happened: never a plain failure
        raise WriteOutcomeUnknown(
            f"Outlook answered {action} without item results, so it is unclear whether the change "
            "was made; it was not retried. Check the mailbox before trying again."
        )
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


def _address(address: str) -> dict[str, str]:
    return {"__type": "EmailAddress:#Exchange", "EmailAddress": address, "RoutingType": "SMTP"}


def _body(text: str) -> dict[str, str]:
    return {"__type": "BodyContentType:#Exchange", "BodyType": "Text", "Value": text}


def _created_id(items: list[dict[str, Any]]) -> str | None:
    inner = items[0].get("Items")
    first = inner[0] if isinstance(inner, list) and inner and isinstance(inner[0], dict) else {}
    item_id = first.get("ItemId")
    value = item_id.get("Id") if isinstance(item_id, dict) else None
    return ids.to_graph(value) if isinstance(value, str) and value else None


class OwsMailWriter:
    """Implements ``MailWriter`` over OWS."""

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


def _create(message: EmailProposal, disposition: str) -> dict[str, Any]:
    """``CreateItem`` as proven by the 2026-10-01 self-send (research §4.2), with the disposition
    switched: ``SendAndSaveCopy`` sends, ``SaveOnly`` leaves a draft in Drafts.

    Replies use EWS's ``ReplyToItem`` / ``ReplyAllToItem`` response objects, which append the
    quoted original; recipients and subject are passed explicitly, so the result matches the
    proposal. (The reply variant is pending its first live check, V2.)
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
            "NewBodyContent": _body(message.body),
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
