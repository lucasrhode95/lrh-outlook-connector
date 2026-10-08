"""Rule writes: stateless proposal, human confirmation, one write, then fresh read-back."""

from __future__ import annotations

import hashlib
import json
from collections.abc import Callable

from outlook_connector.auth.tokens import Account
from outlook_connector.domain.errors import ConnectorError, InvalidRequest, NotFound, WriteOutcomeUnknown
from outlook_connector.domain.models import Folder, InboxRule, RuleChange, RuleWriteResult
from outlook_connector.remote.ports import RuleWriter
from outlook_connector.service.mailbox import Mailbox
from outlook_connector.service.writes import _addresses


class Rules:
    def __init__(
        self, mailbox: Mailbox, writer: RuleWriter, account: Account, check_account: Callable[[], None]
    ) -> None:
        self.mailbox, self.writer, self.account, self.check_account = mailbox, writer, account, check_account

    async def list_rules(self) -> list[InboxRule]:
        """Entry point: read current rules through the bound Outlook account; no cached rules."""
        self.check_account()
        return await self.writer.list_rules()

    async def create_rule(self, changes: RuleChange, confirmation: str | None = None) -> RuleWriteResult:
        """Entry point: propose or confirm a new rule, validating supported fields and account."""
        return await self._write("create", changes=changes, confirmation=confirmation)

    async def update_rule(
        self, rule_id: str, changes: RuleChange, confirmation: str | None = None
    ) -> RuleWriteResult:
        """Entry point: propose or confirm a supported rule edit, or a separate enable/disable toggle."""
        return await self._write("update", rule_id=rule_id, changes=changes, confirmation=confirmation)

    async def reorder_rules(self, rule_ids: list[str], confirmation: str | None = None) -> RuleWriteResult:
        """Entry point: validate a complete permutation; unsupported rules prevent rewriting the list."""
        return await self._write("reorder", rule_ids=rule_ids, confirmation=confirmation)

    async def delete_rule(self, rule_id: str, confirmation: str | None = None) -> RuleWriteResult:
        """Entry point: propose or confirm deletion of a supported rule, never mail deletion."""
        return await self._write("delete", rule_id=rule_id, confirmation=confirmation)

    async def _write(
        self,
        action: str,
        *,
        changes: RuleChange | None = None,
        rule_id: str | None = None,
        rule_ids: list[str] | None = None,
        confirmation: str | None = None,
    ) -> RuleWriteResult:
        """Single validator for public rule write entry points; no wire-format knowledge."""
        before = await self.list_rules()
        by_id = {r.id: r for r in before}
        target = None
        if rule_id is not None:
            if not rule_id.strip():
                raise InvalidRequest("rule_id must not be empty.")
            target = by_id.get(rule_id)
            if target is None:
                raise NotFound("The inbox rule was not found.")
            if target.read_only:
                raise InvalidRequest("This unsupported rule is read-only; manage it in Outlook.")
        folder = None
        if changes is not None:
            changes, folder = await self._validate(changes, creating=action == "create")
        if action == "reorder":
            if not rule_ids or len(rule_ids) != len(set(rule_ids)) or set(rule_ids) != set(by_id):
                raise InvalidRequest("rule_ids must contain every current rule exactly once.")
            if any(r.read_only for r in before):
                raise InvalidRequest(
                    "Cannot reorder while unsupported rules are present: they must remain read-only."
                )
        material = {
            "account": self.account.fingerprint,
            "action": action,
            "rule_id": rule_id,
            "rule_ids": rule_ids,
            "changes": changes.model_dump(exclude_unset=True) if changes else None,
            "state": [(r.id, r.revision) for r in before],
        }
        code = (
            "RULE-" + hashlib.sha256(json.dumps(material, sort_keys=True).encode()).hexdigest()[:16].upper()
        )
        result = RuleWriteResult(
            action=action,
            status="proposed",
            confirmation=code,
            changes=changes,
            rule_id=rule_id,
            rule_ids=rule_ids or [],
            rules=before,
            detail="Review the proposed persistent rule change and confirm this code explicitly.",
        )
        if confirmation is None:
            return result
        if confirmation.strip().upper() != code:
            raise InvalidRequest(
                "Confirmation does not match the current rules and proposed change; propose again."
            )
        self.check_account()
        ambiguous = False
        try:
            if action == "create":
                assert changes is not None
                result.rule_id = await self.writer.create_rule(changes, folder)
            elif action == "update":
                assert changes is not None and target is not None
                await self.writer.update_rule(target, changes, folder)
            elif action == "reorder":
                await self.writer.reorder_rules([by_id[rid] for rid in rule_ids or []])
            else:
                assert target is not None
                await self.writer.delete_rule(target)
        except WriteOutcomeUnknown:
            ambiguous = True
        except ConnectorError as exc:
            result.status, result.confirmation, result.detail = "failed", None, str(exc)
            return result
        try:
            after = await self.writer.list_rules()
        except ConnectorError as exc:
            result.status, result.confirmation = "unknown", None
            result.detail = f"Write was not retried; read-back failed: {exc}. Check Outlook before repeating."
            return result
        result.rules, result.confirmation = after, None
        # The identity NewInboxRule reports can differ from fresh read-back (live 2026-10-07): find the
        # one new rule that matches the request instead.
        if action == "create" and result.rule_id not in {r.id for r in after}:
            added = [r for r in after if r.id not in by_id and _matches(r, changes, folder)]
            result.rule_id = added[0].id if len(added) == 1 else None
        current = next((r for r in after if r.id == result.rule_id), None)
        if action == "delete":
            verified = current is None
        elif action == "reorder":
            verified = [r.id for r in after] == rule_ids and all(
                r.enabled == by_id[r.id].enabled for r in after
            )
        else:
            verified = current is not None and _matches(current, changes, folder)
        result.status = "done" if verified else "unknown" if ambiguous else "failed"
        result.detail = (
            "One write, verified by fresh rule read-back."
            if verified
            else (
                "Read-back did not confirm the requested state; check Outlook before repeating. "
                "No write was retried."
            )
        )
        return result

    async def _validate(self, changes: RuleChange, *, creating: bool) -> tuple[RuleChange, Folder | None]:
        """Assumes (not re-checked here): called only by _write; validates supported rule input once."""
        fields = changes.model_dump(exclude_unset=True)
        if not fields:
            raise InvalidRequest("Supply at least one supported rule field.")
        if creating and not (changes.name or "").strip():
            raise InvalidRequest("A new rule needs a name.")
        if "enabled" in fields and (creating or len(fields) != 1 or changes.enabled is None):
            raise InvalidRequest("Enable/disable must be a separate update, to send each change once.")
        for key in ("name", "stop_processing", "move_to_folder"):
            if key in fields and fields[key] is None:
                raise InvalidRequest(f"{key} cannot be null.")
        if "name" in fields:
            if not str(fields["name"]).strip():
                raise InvalidRequest("Rule name must not be empty.")
            fields["name"] = str(fields["name"]).strip()
        for key in ("from_addresses", "sent_to"):
            if fields.get(key) is not None:
                fields[key] = _addresses(fields[key], key) or None
        for key in ("subject_contains", "subject_or_body_contains"):
            if fields.get(key) is not None:
                if any(not word.strip() for word in fields[key]):
                    raise InvalidRequest(f"{key} words must not be empty.")
                fields[key] = [word.strip() for word in fields[key]] or None
        folder = (
            await self.mailbox.resolve_folder(fields["move_to_folder"])
            if "move_to_folder" in fields
            else None
        )
        if folder:
            fields["move_to_folder"] = folder.id
        return RuleChange.model_validate(fields), folder


def _matches(rule: InboxRule, changes: RuleChange | None, folder: Folder | None) -> bool:
    """Assumes (not re-checked here): validated changes, mapped current rule and resolved folder."""
    assert changes is not None
    for key, expected in changes.model_dump(exclude_unset=True).items():
        if key == "enabled":
            actual = rule.enabled
        elif key == "name":
            actual = rule.name
        elif key == "move_to_folder":
            if not folder or rule.move_to_folder_name != folder.name:
                return False
            continue
        else:
            actual = rule.conditions.model_dump()[key]
        if key in ("from_addresses", "sent_to"):
            addresses = rule.conditions.from_addresses if key == "from_addresses" else rule.conditions.sent_to
            actual = [a.lower() for a in addresses] if addresses else None
            expected = [a.lower() for a in expected] if expected else None
        if actual != expected:
            return False
    return True
