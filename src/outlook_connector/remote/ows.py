"""Outlook Web's JSON RPC (OWS, ``/owa/service.svc``): the write path (architecture §5.5).

A gap fill: Graph cannot write mail with the usable first-party clients (research §2), so every
write goes here with the ``write`` token. Only the ``ows*`` modules know OWS JSON. Contracts are the ones
proven in research §4.1–4.2; anything else is marked where it is used.

- Bearer only: no cookies, no canary. Payloads up to 2,048 URL-encoded characters travel in the
  ``X-OWA-UrlPostData`` header with an empty body, larger ones in the body.
- Every write is sent once. A missing or server-error answer raises ``WriteOutcomeUnknown``.
- Item results carry ``ResponseClass`` / ``ResponseCode``; anything but success is an error that
  names the code.
- The inbox-rule actions use a second style (research §4.4, ``Ows.call_request``): the request object
  is posted as is, and the answer reports ``WasSuccessful`` / ``ErrorCode``.

This module is the client (envelope, headers, results); ``ows_mapping`` builds request bodies and
``ows_mail`` implements ``MailWriter`` with them.
"""

from __future__ import annotations

import json
import re
import uuid
from typing import Any, Protocol
from urllib.parse import quote

from outlook_connector.domain.errors import NotFound, Upstream, WriteOutcomeUnknown
from outlook_connector.remote.transport import Transport

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
