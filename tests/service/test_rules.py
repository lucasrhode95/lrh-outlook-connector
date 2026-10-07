from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from outlook_connector.domain.errors import InvalidRequest
from outlook_connector.domain.models import RuleChange
from outlook_connector.remote.ows_rules import OwsRules
from outlook_connector.service.rules import Rules
from tests.fakes.graph_fake import FakeGraph, sample_mailbox
from tests.service.test_writes import ME, make_writes


@pytest.fixture
def fake() -> FakeGraph:
    return sample_mailbox()


@pytest.fixture
def rules(fake: FakeGraph, tmp_path: Path) -> Rules:
    writes = make_writes(fake, tmp_path)
    return Rules(writes.mailbox, OwsRules(writes.writer._ows), ME, writes.check_account)  # type: ignore[attr-defined]


def changes(**values: Any) -> RuleChange:
    return RuleChange.model_validate(
        {
            "name": "Synthetic rule",
            "from_addresses": ["a@example.com"],
            "sent_to": ["b@example.com"],
            "subject_contains": ["subject"],
            "subject_or_body_contains": ["body"],
            "move_to_folder": "archive",
            "stop_processing": True,
        }
        | values
    )


async def create(rules: Rules) -> str:
    proposed = await rules.create_rule(changes())
    result = await rules.create_rule(changes(), proposed.confirmation)
    assert result.status == "done" and result.rule_id
    return result.rule_id


async def test_proposal_changes_nothing_and_confirmation_is_required(rules: Rules, fake: FakeGraph) -> None:
    proposed = await rules.create_rule(changes())
    assert proposed.status == "proposed" and proposed.confirmation and not fake.inbox_rules
    assert [a for a, _ in fake.ows_calls] == ["GetInboxRule"]
    with pytest.raises(InvalidRequest, match="Confirmation"):
        await rules.create_rule(changes(name="Different"), proposed.confirmation)
    assert not fake.inbox_rules
    result = await rules.create_rule(changes(), proposed.confirmation)
    assert result.status == "done" and len(fake.inbox_rules) == 1
    actions = [a for a, _ in fake.ows_calls]
    assert actions.count("NewInboxRule") == 1 and actions[-1] == "GetInboxRule"
    sent = next(b for a, b in fake.ows_calls if a == "NewInboxRule")["InboxRule"]
    assert sent["MoveToFolder"]["RawIdentity"] == "f-archive"
    assert sent["From"][0]["__type"] == "PeopleIdentity:#Exchange"


async def test_partial_edit_clear_toggle_reorder_delete(rules: Rules, fake: FakeGraph) -> None:
    rid = await create(rules)
    update = RuleChange(name="Renamed", subject_contains=None)
    proposal = await rules.update_rule(rid, update)
    result = await rules.update_rule(rid, update, proposal.confirmation)
    assert result.status == "done" and result.rules[0].conditions.subject_contains is None
    assert result.rules[0].conditions.sent_to == ["b@example.com"]
    assert result.rules[0].move_to_folder_name == "Archive"
    request = next(b for a, b in fake.ows_calls if a == "SetInboxRule")
    assert request["InboxRule"]["SubjectContainsWords"] is None and request["Force"] is False
    for enabled in (False, True):
        update = RuleChange(enabled=enabled)
        proposal = await rules.update_rule(rid, update)
        result = await rules.update_rule(rid, update, proposal.confirmation)
        assert result.status == "done" and result.rules[0].enabled is enabled
    second = await create(rules)
    proposal = await rules.reorder_rules([rid, second])
    result = await rules.reorder_rules([rid, second], proposal.confirmation)
    assert result.status == "done" and [r.id for r in result.rules] == [rid, second]
    proposal = await rules.delete_rule(rid)
    result = await rules.delete_rule(rid, proposal.confirmation)
    assert result.status == "done" and [r.id for r in result.rules] == [second]


@pytest.mark.parametrize(
    "unsupported",
    [
        {"MarkAsRead": True},
        {"ExceptIfSubjectContainsWords": ["private"]},
        {"ForwardTo": [{"SmtpAddress": "x@example.com"}]},
        {"FutureAction": "Default"},
        {"MarkImportance": "High"},  # an active value, not the inactive NullImportance
        {"ExceptIfWithSensitivity": "Private"},
        {"FlaggedForAction": "NullImportance"},  # another field's null value is not this one's
    ],
)
async def test_unsupported_rules_never_rewritten(
    rules: Rules, fake: FakeGraph, unsupported: dict[str, Any]
) -> None:
    rid = await create(rules)
    fake.inbox_rules[0].update(unsupported)
    listed = await rules.list_rules()
    assert listed[0].read_only and listed[0].unsupported
    fake.ows_calls.clear()
    for action in (
        rules.update_rule(rid, RuleChange(name="x")),
        rules.delete_rule(rid),
        rules.reorder_rules([rid]),
    ):
        with pytest.raises(InvalidRequest, match="read-only"):
            await action
    assert all(a == "GetInboxRule" for a, _ in fake.ows_calls)


async def test_changed_server_state_invalidates_confirmation(rules: Rules, fake: FakeGraph) -> None:
    rid = await create(rules)
    update = RuleChange(name="x")
    proposal = await rules.update_rule(rid, update)
    fake.inbox_rules[0]["Name"] = "Changed externally"
    fake.ows_calls.clear()
    with pytest.raises(InvalidRequest, match="Confirmation"):
        await rules.update_rule(rid, update, proposal.confirmation)
    assert [a for a, _ in fake.ows_calls] == ["GetInboxRule"]


@pytest.mark.parametrize(
    "script, expected", [("done-no-answer", "done"), ("no-answer", "unknown"), (429, "failed")]
)
async def test_single_write_outcomes(rules: Rules, fake: FakeGraph, script: Any, expected: str) -> None:
    proposal = await rules.create_rule(changes())
    fake.ows_calls.clear()
    fake.ows_next = [None, script]
    result = await rules.create_rule(changes(), proposal.confirmation)
    assert result.status == expected
    assert [a for a, _ in fake.ows_calls].count("NewInboxRule") == 1


async def test_failed_readback_keeps_unknown(rules: Rules, fake: FakeGraph) -> None:
    proposal = await rules.create_rule(changes())
    fake.ows_next = [None, None, 503]
    result = await rules.create_rule(changes(), proposal.confirmation)
    assert result.status == "unknown" and result.rule_id and len(fake.inbox_rules) == 1


@pytest.mark.parametrize(
    "values", [{"name": ""}, {"from_addresses": ["bad"]}, {"subject_contains": [""]}, {"enabled": False}]
)
async def test_invalid_create_never_writes(rules: Rules, fake: FakeGraph, values: dict[str, Any]) -> None:
    with pytest.raises(InvalidRequest):
        await rules.create_rule(changes(**values))
    assert all(a == "GetInboxRule" for a, _ in fake.ows_calls)


async def test_invalid_reorder_and_combined_toggle(rules: Rules) -> None:
    rid = await create(rules)
    for ids in ([], [rid, rid], ["wrong"]):
        with pytest.raises(InvalidRequest):
            await rules.reorder_rules(ids)
    with pytest.raises(InvalidRequest, match="separate"):
        await rules.update_rule(rid, RuleChange(enabled=False, name="x"))


async def test_account_mismatch_blocks_rule_reads_and_writes(fake: FakeGraph, tmp_path: Path) -> None:
    from outlook_connector.auth.tokens import Account
    from outlook_connector.domain.errors import AccountMismatch

    writes = make_writes(
        fake, tmp_path, account=Account(tenant_id="tenant-x", object_id="other", username="me@example.com")
    )
    rules = Rules(writes.mailbox, OwsRules(writes.writer._ows), writes.account, writes.check_account)  # type: ignore[attr-defined]
    with pytest.raises(AccountMismatch):
        await rules.list_rules()
    with pytest.raises(AccountMismatch):
        await rules.create_rule(changes())
    assert not fake.ows_calls


async def test_readback_mismatch_reports_failure_without_retry(rules: Rules, fake: FakeGraph) -> None:
    rid = await create(rules)
    update = RuleChange(name="x")
    proposal = await rules.update_rule(rid, update)
    fake.ows_next = [None, {"WasSuccessful": True, "ErrorCode": 0}]
    result = await rules.update_rule(rid, update, proposal.confirmation)
    assert result.status == "failed" and result.rules[0].name == "Synthetic rule"


async def test_reorder_preserves_disabled_rules(rules: Rules, fake: FakeGraph) -> None:
    first = await create(rules)
    second = await create(rules)
    toggle = RuleChange(enabled=False)
    proposal = await rules.update_rule(first, toggle)
    await rules.update_rule(first, toggle, proposal.confirmation)
    proposal = await rules.reorder_rules([first, second])
    result = await rules.reorder_rules([first, second], proposal.confirmation)
    assert result.status == "done" and not result.rules[0].enabled


async def test_proposal_shows_only_supplied_fields(rules: Rules, fake: FakeGraph) -> None:
    created = await rules.create_rule(changes())
    await rules.create_rule(changes(), created.confirmation)
    rule_id = (await rules.list_rules())[0].id
    proposal = await rules.update_rule(rule_id, RuleChange(name="Renamed", subject_contains=None))
    shown = proposal.model_dump(mode="json")["changes"]
    assert shown == {"name": "Renamed", "subject_contains": None}  # omitted fields are not "clear"
    replayed = RuleChange.model_validate(shown)  # a later call can resend exactly what was shown
    result = await rules.update_rule(rule_id, replayed, proposal.confirmation)
    assert result.status == "done"


async def test_live_rule_shapes_stay_supported_and_creation_id_is_reconciled(
    rules: Rules, fake: FakeGraph
) -> None:
    """Live 2026-10-07: every rule carries description metadata and inactive enum values, and
    NewInboxRule reports an identity that fresh GetInboxRule does not (research §4.5)."""
    rid = await create(rules)
    (listed,) = await rules.list_rules()
    assert listed.id == rid == fake.inbox_rules[0]["Identity"]["RawIdentity"]
    assert not listed.read_only and listed.unsupported == []
    assert fake.inbox_rules[0]["MarkImportance"] == "NullImportance"  # the shape was present
    result = await rules.update_rule(rid, RuleChange(subject_contains=None))
    done = await rules.update_rule(rid, RuleChange(subject_contains=None), result.confirmation)
    assert done.status == "done"
