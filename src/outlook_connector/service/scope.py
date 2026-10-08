"""Validation shared by public service entry points that accept mailbox scope."""

from outlook_connector.domain.errors import InvalidRequest
from outlook_connector.domain.models import Scope


def validate_scope(
    scope: Scope,
    *,
    sent_items: bool = True,
    meeting_mail: bool = True,
    deleted_items: bool = True,
) -> None:
    """Reject non-default scope settings that the current operation cannot apply.

    Scope defaults are accepted even when their keys are irrelevant. A public entry point declares
    which keys its operation can honor; internal helpers trust that validated contract.
    """
    if not sent_items and not scope.sent_items:
        raise InvalidRequest("scope.sent_items has no effect for this operation.")
    if not meeting_mail and not scope.meeting_mail:
        raise InvalidRequest("scope.meeting_mail has no effect for this operation.")
    if not deleted_items and scope.deleted_items:
        raise InvalidRequest("scope.deleted_items has no effect for this operation.")
