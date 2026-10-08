"""Native roaming signatures over Outlook Cloud Settings (research §4.7–4.8)."""

from __future__ import annotations

import hashlib
import json
import time
from typing import Any
from urllib.parse import urlencode

from outlook_connector.auth.tokens import TokenProvider
from outlook_connector.domain.errors import AccountMismatch, Upstream
from outlook_connector.remote.ports import SignatureContents, SignatureSettings
from outlook_connector.remote.transport import Transport, operation

BASE_URL = "https://outlook.cloud.microsoft/ows/v1/OutlookCloudSettings/settings/"
PROFILE = "outlook"
LIST_SETTING = "roaming_signature_list"
NEW_SETTING = "roaming_new_signature"
REPLY_SETTING = "roaming_reply_signature"


class CloudSettings:
    """Native signature settings adapter. All wire names and payloads stay in this module."""

    def __init__(self, transport: Transport, tokens: TokenProvider) -> None:
        self._transport = transport
        self._tokens = tokens

    def account(self) -> dict[str, Any]:
        """Identity of the Outlook profile used for Cloud Settings."""
        return self._tokens.get_token(PROFILE).claims()

    async def settings(self) -> SignatureSettings:
        """Read fresh settings; absent records mean no names/defaults, not an upstream failure.

        Present records must still have one consistent scope. An entirely empty response has no
        write scope; reads may report no signature, but writes must obtain a scope from Outlook.
        """
        query = urlencode({"settingname": f"{LIST_SETTING},{NEW_SETTING},{REPLY_SETTING}"})
        with operation("reading native signatures"):
            data = await self._transport.json(
                "GET",
                f"{BASE_URL}?{query}",
                profile=PROFILE,
                headers=_headers(),
            )
        records = _records(data)
        listed = _one(records, LIST_SETTING)
        new = _one(records, NEW_SETTING)
        reply = _one(records, REPLY_SETTING)
        present = tuple(record for record in (listed, new, reply) if record is not None)
        if len(present) != len(records):
            raise Upstream("Outlook returned unexpected native signature settings.")
        scope = present[0].get("scope") if present else None
        if present and any(record.get("scope") is None for record in present):
            raise Upstream("Outlook did not return the account scope for native signatures.")
        if any(record.get("scope") != scope for record in present):
            raise AccountMismatch("Native signature settings do not share the bound account scope.")
        raw_names = listed.get("value") if listed is not None else ""
        if not isinstance(raw_names, str):
            raise Upstream("Outlook returned an unreadable native signature list.")
        names = tuple(name for name in raw_names.split(",") if name)
        new_default = _default(new)
        reply_default = _default(reply)
        revision = _revision(present)
        return SignatureSettings(names, new_default, reply_default, scope, revision)

    async def contents(self, name: str, settings: SignatureSettings) -> SignatureContents | None:
        """Read the formats for a name from a fresh settings snapshot."""
        query = urlencode({"settingname": name})
        with operation("reading native signature contents"):
            data = await self._transport.json(
                "GET",
                f"{BASE_URL}?{query}",
                profile=PROFILE,
                headers=_headers(large=True),
            )
        records = _records(data)
        if not records:
            return None
        values: dict[str, str] = {}
        for record in records:
            if record.get("scope") != settings.scope:
                raise AccountMismatch("Native signature contents do not match the bound account scope.")
            if record.get("name") != name:
                raise Upstream("Outlook returned contents for a different native signature.")
            key, value = record.get("secondaryKey"), record.get("value")
            if key in {"htm", "txt"} and isinstance(value, str):
                values[key] = value
        return SignatureContents(values.get("htm"), values.get("txt"), _revision(tuple(records)))

    async def create(self, name: str, html: str, text: str, settings: SignatureSettings) -> None:
        """Create one signature in a single Cloud Settings write."""
        await self._write_contents(name, html, text, settings)

    async def update(self, name: str, html: str, text: str, settings: SignatureSettings) -> None:
        """Replace one signature's contents in a single Cloud Settings write."""
        await self._write_contents(name, html, text, settings)

    async def _write_contents(self, name: str, html: str, text: str, settings: SignatureSettings) -> None:
        _require_write_scope(settings)
        timestamp = 621355968000000000 + time.time_ns() // 100
        payload = [
            _content_record(name, settings.scope, "htm", html, timestamp),
            _content_record(name, settings.scope, "txt", text, timestamp),
        ]
        with operation("saving a native signature"):
            await self._transport.json(
                "PATCH",
                f"{BASE_URL}account",
                profile=PROFILE,
                headers=_headers(large=True, override_timestamp=False),
                json_body=payload,
                write=True,
            )

    async def delete(self, name: str) -> None:
        """Delete all native formats for one name in a single write."""
        with operation("deleting a native signature"):
            await self._transport.json(
                "DELETE",
                f"{BASE_URL}account",
                profile=PROFILE,
                headers={
                    **_headers(large=True),
                    "Content-Type": "application/json",
                },
                json_body={"name": name},
                write=True,
            )

    async def set_default(self, name: str | None, which: str, settings: SignatureSettings) -> None:
        """Set one or both Outlook defaults in one write; None clears the pointer."""
        _require_write_scope(settings)
        setting_names = (
            [NEW_SETTING, REPLY_SETTING]
            if which == "both"
            else [NEW_SETTING if which == "new" else REPLY_SETTING]
        )
        payload = [
            {
                "itemClass": "RoamingSetting",
                "name": setting_name,
                "scope": settings.scope,
                "secondaryKey": setting_name,
                "type": "String",
                "value": name or "",
            }
            for setting_name in setting_names
        ]
        with operation("setting native signature defaults"):
            await self._transport.json(
                "PATCH",
                f"{BASE_URL}account",
                profile=PROFILE,
                headers=_headers(large=False, override_timestamp=False),
                json_body=payload,
                write=True,
            )


def _headers(*, large: bool | None = None, override_timestamp: bool | None = None) -> dict[str, str]:
    headers = {"x-outlook-client": "owa", "Accept": "application/json"}
    if large is not None:
        headers["x-islargesetting"] = str(large).lower()
    if override_timestamp is not None:
        headers["x-overridetimestamp"] = str(override_timestamp).lower()
    return headers


def _records(data: Any) -> list[dict[str, Any]]:
    if isinstance(data, list) and all(isinstance(item, dict) for item in data):
        return data
    if isinstance(data, dict) and isinstance(data.get("value"), list):
        values = data["value"]
        if all(isinstance(item, dict) for item in values):
            return values
    raise Upstream("Outlook returned an unreadable native signature settings response.")


def _one(records: list[dict[str, Any]], name: str) -> dict[str, Any] | None:
    matches = [record for record in records if record.get("name") == name]
    if len(matches) > 1:
        raise Upstream("Outlook returned duplicate native signature settings.")
    return matches[0] if matches else None


def _default(record: dict[str, Any] | None) -> str | None:
    if record is None:
        return None
    value = record.get("value")
    if value is None or value == "":
        return None
    if not isinstance(value, str):
        raise Upstream("Outlook returned an unreadable native signature default.")
    return value


def _require_write_scope(settings: SignatureSettings) -> None:
    if settings.scope is None:
        raise Upstream(
            "Outlook did not return an account scope for signature writes. "
            "Configure a native signature in Outlook, then retry."
        )


def _revision(records: tuple[dict[str, Any], ...]) -> str:
    state = [
        {
            key: record.get(key)
            for key in ("name", "value", "scope", "Timestamp", "timestamp", "secondaryKey")
            if key in record
        }
        for record in records
    ]
    encoded = json.dumps(state, ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()


def _content_record(name: str, scope: Any, key: str, value: str, timestamp: int) -> dict[str, Any]:
    return {
        "itemClass": "RoamingSetting",
        "name": name,
        "scope": scope,
        "secondaryKey": key,
        "type": "Blob",
        "value": value,
        "parentSetting": LIST_SETTING,
        "metadata": "encoding:utf-8",
        "timestamp": timestamp,
        "value@is.Large": True,
    }
