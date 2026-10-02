"""Self-contained continuation cursors.

Processes are short-lived (architecture §2), so a cursor carries everything needed to continue:
the remote continuation link and the original selection. Cursors are opaque to callers.
"""

from __future__ import annotations

import base64
import json
from typing import Any

from outlook_connector.domain.errors import InvalidRequest


def encode(kind: str, **state: Any) -> str:
    raw = json.dumps({"kind": kind, **state}, separators=(",", ":"), default=str).encode()
    return base64.urlsafe_b64encode(raw).decode().rstrip("=")


def decode(cursor: str, kind: str) -> dict[str, Any]:
    try:
        data = json.loads(base64.urlsafe_b64decode(cursor + "=" * (-len(cursor) % 4)))
    except (ValueError, json.JSONDecodeError):
        raise InvalidRequest("Malformed cursor.") from None
    if not isinstance(data, dict) or data.get("kind") != kind:
        raise InvalidRequest(f"This cursor does not belong to {kind}.")
    return data
