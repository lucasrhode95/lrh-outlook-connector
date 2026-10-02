"""Command-line entry point: ``outlook-connector <command>``.

Commands import their dependencies lazily, so that ``mcp`` never loads the web stack.
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import UTC, datetime
from typing import TYPE_CHECKING

from outlook_connector import config
from outlook_connector.domain.errors import AuthenticationRequired, ConnectorError

if TYPE_CHECKING:
    from outlook_connector.auth.tokens import AccessToken, CacheStatus, TokenProvider

EXIT_OK = 0
EXIT_ERROR = 1
EXIT_AUTH_REQUIRED = 2
EXIT_INTERRUPTED = 130


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog=config.CLI_NAME, description="Local Outlook mailbox connector.")
    sub = parser.add_subparsers(dest="command", required=True)

    auth = sub.add_parser("auth", help="Sign in with a device code (the only command that ever prompts).")
    auth.add_argument(
        "profile",
        nargs="?",
        default="read",
        choices=sorted(config.PROFILES),
        help="Which client profile to sign in (default: read).",
    )
    auth.add_argument("--force", action="store_true", help="Sign in again even if a silent token works.")
    auth.add_argument("--sign-out", action="store_true", help="Delete the token cache (all profiles).")
    _add_unsecure(auth)

    status = sub.add_parser("status", help="Show the signed-in account and profiles (offline by default).")
    status.add_argument(
        "--check",
        action="store_true",
        help="Also acquire each profile's token silently (contacts Microsoft).",
    )
    status.add_argument("--json", action="store_true", help="Machine-readable output.")
    _add_unsecure(status)

    mcp = sub.add_parser("mcp", help="Run the MCP server over stdio (started by your MCP client).")
    _add_unsecure(mcp)
    return parser


def _add_unsecure(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--unsecure", action="store_true", help="Development only: use a separate PLAINTEXT token cache."
    )


def main(argv: list[str] | None = None, *, provider_factory: type[TokenProvider] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.unsecure:
        print(
            "WARNING: --unsecure stores reusable Microsoft tokens as PLAINTEXT at "
            f"{config.token_cache_path(unsecure=True)}. Do not share it; delete it when done.",
            file=sys.stderr,
        )
    if args.command == "mcp":
        from outlook_connector.surfaces import mcp_main

        mcp_main.run(unsecure=args.unsecure)  # stdout belongs to the MCP protocol from here on
        return EXIT_OK
    try:
        if provider_factory is None:
            from outlook_connector.auth.tokens import TokenProvider as provider_factory
        provider = provider_factory(unsecure=args.unsecure)
        if args.command == "auth":
            return _auth(provider, args)
        return _status(provider, args)
    except AuthenticationRequired as exc:
        print(str(exc), file=sys.stderr)
        return EXIT_AUTH_REQUIRED
    except ConnectorError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return EXIT_ERROR
    except KeyboardInterrupt:
        print("Interrupted.", file=sys.stderr)
        return EXIT_INTERRUPTED


def _auth(provider: TokenProvider, args: argparse.Namespace) -> int:
    if args.sign_out:
        existed = provider.sign_out()
        print("Signed out: token cache deleted." if existed else "No token cache to delete.")
        return EXIT_OK
    profile = config.PROFILES[args.profile]
    if not args.force:
        try:
            token = provider.get_token(args.profile)
        except AuthenticationRequired:
            pass
        else:
            print(f"Already signed in for '{args.profile}' ({profile.purpose}).")
            _print_token_summary(token)
            return EXIT_OK
    print(f"Signing in for '{args.profile}': {profile.purpose}.")
    token = provider.sign_in(args.profile, show=lambda message: print(f"\n{message}\n", flush=True))
    print(f"Signed in for '{args.profile}'.")
    _print_token_summary(token)
    return EXIT_OK


def _status(provider: TokenProvider, args: argparse.Namespace) -> int:
    status = provider.status()
    checks: dict[str, dict[str, object]] = {}
    if args.check:
        for name in config.PROFILES:
            try:
                checks[name] = {"ok": True, **_token_facts(provider.get_token(name))}
            except ConnectorError as exc:
                checks[name] = {"ok": False, "error": str(exc)}
    if args.json:
        print(json.dumps(_status_dict(status, checks), indent=2))
    else:
        _print_status(status, checks)
    signed_in_everywhere_checked = all(c["ok"] for c in checks.values()) if checks else True
    return EXIT_OK if signed_in_everywhere_checked else EXIT_AUTH_REQUIRED


# ---------------------------------------------------------------------- formatting


def _token_facts(token: AccessToken) -> dict[str, object]:
    claims = token.claims()
    scopes = str(claims.get("scp", "")).split()
    return {
        "source": token.source,
        "audience": claims.get("aud"),
        "account": claims.get("upn") or claims.get("unique_name"),
        "scope_count": len(scopes),
        "mail_scopes": sorted(s for s in scopes if s.startswith("Mail.")),
        "expires_on": _iso(token.expires_on),
    }


def _print_token_summary(token: AccessToken) -> None:
    facts = _token_facts(token)
    print(f"  account:     {facts['account']}")
    print(f"  audience:    {facts['audience']}")
    print(
        f"  mail scopes: {', '.join(facts['mail_scopes']) or '(none)'}  "
        f"({facts['scope_count']} scopes in total)"
    )
    print(f"  token from:  {facts['source']}, expires {facts['expires_on']}")


def _status_dict(status: CacheStatus, checks: dict[str, dict[str, object]]) -> dict[str, object]:
    return {
        "cache": {"mode": status.mode, "path": str(status.path), "exists": status.exists},
        "accounts": [{"username": a.username, "fingerprint": a.fingerprint} for a in status.accounts],
        "profiles": [
            {
                "profile": p.profile,
                "purpose": p.purpose,
                "client_id": p.client_id,
                "scopes": list(p.scopes),
                "signed_in": p.signed_in,
                "access_token_expires_on": _iso(p.access_token_expires_on),
                **({"check": checks[p.profile]} if p.profile in checks else {}),
            }
            for p in status.profiles
        ],
    }


def _print_status(status: CacheStatus, checks: dict[str, dict[str, object]]) -> None:
    print(f"Token cache: {status.mode}, {status.path}" + ("" if status.exists else " (not created yet)"))
    if not status.accounts:
        print("Account:     not signed in")
    for account in status.accounts:
        print(f"Account:     {account.username} (fingerprint {account.fingerprint})")
    for p in status.profiles:
        state = "signed in" if p.signed_in else "not signed in"
        print(f"\n[{p.profile}] {p.purpose}")
        print(f"  client {p.client_id} -> {' '.join(p.scopes)}")
        print(f"  {state}; cached access token valid until {_iso(p.access_token_expires_on) or '-'}")
        if p.profile in checks:
            check = checks[p.profile]
            if check["ok"]:
                print(
                    f"  check: ok ({check['source']}), "
                    f"mail scopes {', '.join(check['mail_scopes']) or '(none)'}"
                )
            else:
                print(f"  check: FAILED - {check['error']}")


def _iso(epoch: int | None) -> str | None:
    return datetime.fromtimestamp(epoch, UTC).isoformat(timespec="seconds") if epoch else None


if __name__ == "__main__":
    raise SystemExit(main())
