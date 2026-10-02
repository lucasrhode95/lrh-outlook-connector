"""Which Graph mail scopes can our first-party clients obtain? (token requests only, no mailbox calls)

Uses each client's stored refresh token to request additional Graph scopes. Recorded
AADSTS65002 denials are skipped. Output: granted/denied, audience, mail scope names.
"""

from __future__ import annotations

from common import READ_CLIENT, WRITE_CLIENT, emit, try_scope

GRAPH = "https://graph.microsoft.com/"
CASES = {
    "outlook_mobile": (READ_CLIENT, ["Mail.ReadWrite", "Mail.Send", "Mail.ReadWrite.Shared", ".default"]),
    "one_outlook_web": (WRITE_CLIENT, ["Mail.Read", "Mail.ReadWrite", "Mail.Send", ".default"]),
}


def main() -> int:
    emit({name: {s: try_scope(client, GRAPH + s) for s in scopes} for name, (client, scopes) in CASES.items()})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
