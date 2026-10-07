"""Inbox-rule contracts proven in API research §4.4; only this adapter knows OWS fields."""

from __future__ import annotations

import hashlib
import json
from typing import Any

from outlook_connector.domain.errors import Upstream
from outlook_connector.domain.models import Folder, InboxRule, RuleChange
from outlook_connector.remote.ows import SERVER_VERSION, Ows
from outlook_connector.remote.ports import RuleWriter

FIELDS = {
    "name": "Name",
    "from_addresses": "From",
    "sent_to": "SentTo",
    "subject_contains": "SubjectContainsWords",
    "subject_or_body_contains": "SubjectOrBodyContainsWords",
    "move_to_folder": "MoveToFolder",
    "stop_processing": "StopProcessingRules",
}
METADATA = {
    "__type",
    "Identity",
    "Name",
    "Enabled",
    "Priority",
    "Description",
    "SupportedByTask",
    "RuleProvider",
    "InError",
    "ErrorType",
    "IsValid",
    "ObjectState",
    "DescriptionTimeFormat",
    "DescriptionTimeZone",
}
# The value OWS reports for each inactive condition/action on every rule (live 2026-10-07, research
# §4.5): this exact value means "not set". Any other value is active behavior the connector does
# not manage, so the rule stays read-only.
INACTIVE = {
    "FlaggedForAction": "NullInboxRuleMessageFlag",
    "RequestedAction": "NullInboxRuleMessageFlag",
    "ExceptIfFlaggedForAction": "NullInboxRuleMessageFlag",
    "ExceptIfRequestedAction": "NullInboxRuleMessageFlag",
    "MessageTypeMatches": "NullInboxRuleMessageType",
    "ExceptIfMessageTypeMatches": "NullInboxRuleMessageType",
    "MarkImportance": "NullImportance",
    "WithImportance": "NullImportance",
    "ExceptIfWithImportance": "NullImportance",
    "WithSensitivity": "NullSensitivity",
    "ExceptIfWithSensitivity": "NullSensitivity",
}


class OwsRules(RuleWriter):
    def __init__(self, ows: Ows) -> None:
        self.ows = ows

    async def list_rules(self) -> list[InboxRule]:
        """Entry point from the server: validate the collection and map all rules, including read-only."""
        answer = await self.ows.call_request("GetInboxRule", {"UseServerRulesLoader": True})
        collection = answer.get("InboxRuleCollection")
        values = collection.get("InboxRules") if isinstance(collection, dict) else None
        if not isinstance(values, list) or any(not isinstance(v, dict) for v in values):
            raise Upstream("Outlook returned an unreadable inbox-rule collection.")
        rules = [_rule(value) for value in values]
        if len({r.id for r in rules}) != len(rules):
            raise Upstream("Outlook returned duplicate inbox-rule identities.")
        return sorted(rules, key=lambda r: r.priority)

    async def create_rule(self, changes: RuleChange, folder: Folder | None) -> str | None:
        """Assumes (not re-checked here): supported, validated, confirmed fields and bound account.

        Returns the identity Outlook reports for the new rule. Live, it can differ from the identity
        fresh ``GetInboxRule`` reports (research §4.5), so the service reconciles it against read-back.
        """
        answer = await self.ows.call_request("NewInboxRule", {"InboxRule": _fields(changes, folder)})
        value = answer.get("InboxRule")
        identity = value.get("Identity") if isinstance(value, dict) else None
        return identity.get("RawIdentity") if isinstance(identity, dict) else None

    async def update_rule(self, rule: InboxRule, changes: RuleChange, folder: Folder | None) -> None:
        """Assumes (not re-checked here): supported existing rule; toggles are separate single writes."""
        if "enabled" in changes.model_fields_set:
            action = "EnableInboxRule" if changes.enabled else "DisableInboxRule"
            await self.ows.call_request(action, {"Identity": _identity(rule)})
        else:
            fields = _fields(changes, folder)
            fields.setdefault("Name", rule.name)
            fields.setdefault("From", _people(rule.conditions.from_addresses))
            fields.setdefault("SentTo", _people(rule.conditions.sent_to))
            fields.setdefault("StopProcessingRules", rule.conditions.stop_processing)
            fields["Identity"] = {"DisplayName": fields["Name"], "RawIdentity": rule.id}
            await self.ows.call_request("SetInboxRule", {"InboxRule": fields, "Force": False})

    async def reorder_rules(self, rules: list[InboxRule]) -> None:
        """Assumes (not re-checked here): complete permutation of supported current rules."""
        header = {"__type": "JsonRequestHeaders:#Exchange", "RequestServerVersion": SERVER_VERSION}
        await self.ows.call_request(
            "SetInboxAndSweepRules",
            {
                "EnableDisableInboxRules": [
                    {
                        "__type": "EnableDisableInboxRuleRequest:#Exchange",
                        "Header": header,
                        "Identity": _identity(rule),
                        "IsEnabled": rule.enabled,
                    }
                    for rule in rules
                ]
            },
        )

    async def delete_rule(self, rule: InboxRule) -> None:
        """Assumes (not re-checked here): supported, confirmed existing rule."""
        await self.ows.call_request("RemoveInboxRule", {"Identity": _identity(rule)})


def _rule(value: dict[str, Any]) -> InboxRule:
    identity = value.get("Identity")
    if not isinstance(identity, dict) or not isinstance(identity.get("RawIdentity"), str):
        raise Upstream("Outlook returned a rule without a readable identity.")
    if not isinstance(value.get("Enabled"), bool) or not isinstance(value.get("Priority"), int):
        raise Upstream("Outlook returned an unreadable rule state or priority.")
    unsupported = [
        key
        for key, v in value.items()
        if key not in METADATA | set(FIELDS.values())
        and v not in (None, False, "", [], {})
        and not (key in ("DisplayAlert", "PlaySound") and v == "Default")
        and not (key in INACTIVE and v == INACTIVE[key])
    ]
    if value.get("InError"):
        unsupported.append("InError")
    conditions: dict[str, Any] = {}
    for public, wire in FIELDS.items():
        v = value.get(wire)
        if public in ("from_addresses", "sent_to"):
            if v is not None and (
                not isinstance(v, list) or any(not isinstance(a, dict) or not a.get("SmtpAddress") for a in v)
            ):
                unsupported.append(wire)
                v = None
            else:
                v = [a["SmtpAddress"] for a in v] if v else None
        if (
            public in ("subject_contains", "subject_or_body_contains")
            and v is not None
            and (not isinstance(v, list) or any(not isinstance(word, str) for word in v))
        ):
            unsupported.append(wire)
            v = None
        if public == "stop_processing" and v is not None and not isinstance(v, bool):
            unsupported.append(wire)
            v = None
        if public != "move_to_folder":
            conditions[public] = v
    folder = value.get("MoveToFolder") or {}
    if not isinstance(folder, dict):
        unsupported.append("MoveToFolder")
        folder = {}
    description = value.get("Description") or []
    return InboxRule(
        id=identity["RawIdentity"],
        name=value.get("Name") or identity.get("DisplayName") or "",
        enabled=value["Enabled"],
        priority=value["Priority"],
        conditions=RuleChange.model_validate(conditions),
        move_to_folder_name=folder.get("DisplayName"),
        move_to_folder_reference=folder.get("RawIdentity"),
        unsupported=sorted(set(unsupported)),
        read_only=bool(unsupported),
        description=description if isinstance(description, list) else [str(description)],
        revision=hashlib.sha256(json.dumps(value, sort_keys=True).encode()).hexdigest(),
    )


def _identity(rule: InboxRule) -> dict[str, str]:
    return {"DisplayName": rule.id, "RawIdentity": rule.id}


def _people(addresses: list[str] | None) -> list[dict[str, str]] | None:
    return (
        [
            {"__type": "PeopleIdentity:#Exchange", "DisplayName": a, "SmtpAddress": a, "RoutingType": "SMTP"}
            for a in addresses
        ]
        if addresses
        else None
    )


def _fields(changes: RuleChange, folder: Folder | None) -> dict[str, Any]:
    fields = {}
    for key, value in changes.model_dump(exclude_unset=True).items():
        if key in ("from_addresses", "sent_to"):
            value = _people(value)
        elif key == "move_to_folder" and folder:
            value = {"DisplayName": folder.name, "RawIdentity": folder.id}
        fields[FIELDS[key]] = value
    return fields
