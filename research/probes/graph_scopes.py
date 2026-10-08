"""Which Graph mail scopes can our first-party clients obtain? (token requests only, no mailbox calls)

Uses each client's stored refresh token to request additional Graph scopes. Recorded
Denials configured for the current environment are skipped; none are skipped by default. Output: granted/denied, audience, mail scope names.
"""

from __future__ import annotations

from common import GRAPH_RESOURCE, GRAPH_CLIENT, OUTLOOK_CLIENT, emit, try_scope

GRAPH = GRAPH_RESOURCE + "/"
CASES = {
    "outlook_mobile": (GRAPH_CLIENT, ["Mail.ReadWrite", "Mail.Send", "Mail.ReadWrite.Shared", ".default"]),
    "one_outlook_web": (OUTLOOK_CLIENT, ["Mail.Read", "Mail.ReadWrite", "Mail.Send", ".default"]),
}


def main() -> int:
    emit({name: {s: try_scope(client, GRAPH + s) for s in scopes} for name, (client, scopes) in CASES.items()})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
