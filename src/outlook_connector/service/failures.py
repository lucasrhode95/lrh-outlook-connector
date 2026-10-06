"""One classification of failed Microsoft requests for exports and conversations (requirements v4 §10.1).

Every gap in an export or a conversation body is an ExportError: the step, Microsoft's answer (status,
code, message, request id), the likely cause, whether retrying can help, and what to do. It is
rendered the same way everywhere: ``error_block`` for text, the model itself for JSONL and MCP.
"""

from __future__ import annotations

from collections import Counter
from collections.abc import Sequence

from outlook_connector.domain.errors import Failure
from outlook_connector.domain.models import ExportError, ExportStep

THROTTLING = (429, 503)
MAX_MESSAGE = 200

# likely cause, retry, fix, short label for summaries
THROTTLED = (
    "Microsoft throttled the mailbox (about 4 parallel requests or 10,000 per 10 minutes); "
    "the message itself is fine",
    True,
    "export it again in a few minutes",
    "throttled",
)
SERVICE = ("Microsoft service or network problem", True, "retry later", "service or network problem")
DENIED = (
    "access denied for this item (for example an encrypted or protected message)",
    False,
    "retrying will not help",
    "access denied",
)
GONE = (
    "deleted or moved in Outlook during the export",
    False,
    "refresh and select it again",
    "deleted or moved",
)
UNEXPECTED = ("unexpected error", False, "report it with the request id", "unexpected error")
LABELS = {cause: label for cause, _, _, label in (THROTTLED, SERVICE, DENIED, GONE, UNEXPECTED)}


def export_error(step: ExportStep, failure: Failure | None, *, message: str | None = None) -> ExportError:
    """Classify a failed request. ``failure`` None: the step failed without a Microsoft answer
    (``message`` says why)."""
    status = failure.status if failure else None
    if failure is None:
        cause = UNEXPECTED
    elif status in THROTTLING:
        cause = THROTTLED
    elif status is None or status >= 500:
        cause = SERVICE
    elif status == 403:
        cause = DENIED
    elif status == 404:
        cause = GONE
    else:
        cause = UNEXPECTED
    likely, retry, fix, _ = cause
    return ExportError(
        step=step,
        status=status,
        code=failure.code if failure else None,
        message=failure.message if failure else (message[:MAX_MESSAGE] if message else None),
        request_id=failure.request_id if failure else None,
        likely_cause=likely,
        retry=retry,
        fix=fix,
    )


def error_from(step: ExportStep, exc: Exception) -> ExportError:
    """Classify a raised error: transport errors carry Microsoft's answer as ``failure``."""
    return export_error(step, getattr(exc, "failure", None), message=str(exc))


def gone(step: ExportStep) -> ExportError:
    return export_error(step, Failure(status=404, code="ErrorItemNotFound"))


def describe(error: ExportError) -> str:
    """'HTTP 429 TooManyRequests: <message>, request-id <id>', or what happened without an answer."""
    if error.status:
        text = f"HTTP {error.status} {error.code or 'no error code'}"
        text += f": {error.message}" if error.message else ""
    else:
        text = error.message or "no response from Microsoft"
    return f"{text}, request-id {error.request_id}" if error.request_id else text


def error_block(headline: str, error: ExportError) -> str:
    """The text marker for a gap, in exports and get_conversation."""
    return "\n".join(
        (
            f"[EXPORT ERROR] {headline}",
            f"  Step:   {error.step}",
            f"  Error:  {describe(error)}",
            f"  Likely: {error.likely_cause}",
            f"  Fix:    {error.fix}",
        )
    )


def error_summary(
    bodies: Sequence[ExportError], attachments: Sequence[ExportError], listings: Sequence[ExportError]
) -> str | None:
    """The export header's one line about everything that could not be exported."""
    parts = []
    if bodies:
        parts.append(f"{len(bodies)} message {'body' if len(bodies) == 1 else 'bodies'} ({_causes(bodies)})")
    if attachments:
        noun = "attachment" if len(attachments) == 1 else "attachments"
        parts.append(f"{len(attachments)} {noun} ({_causes(attachments)})")
    if listings:
        noun = "message" if len(listings) == 1 else "messages"
        parts.append(f"the attachments of {len(listings)} {noun} ({_causes(listings)})")
    if not parts:
        return None
    joined = parts[0] if len(parts) == 1 else ", ".join(parts[:-1]) + " and " + parts[-1]
    return f"Export errors: {joined} could not be exported; they are marked [EXPORT ERROR] below."


def _causes(errors: Sequence[ExportError]) -> str:
    counts = Counter(LABELS.get(e.likely_cause, e.likely_cause) for e in errors)
    return ", ".join(f"{count} {label}" for label, count in counts.most_common())
