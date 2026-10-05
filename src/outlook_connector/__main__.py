"""Command-line entry point: ``outlook-connector <command>``.

Top-down: ``main`` first, then one handler per command, then the parser and the output helpers.
Handlers import their dependencies lazily, so that ``mcp`` never loads the web stack.
"""

from __future__ import annotations

import argparse
import contextlib
import functools
import json
import sys
from collections.abc import Callable
from datetime import UTC, datetime
from typing import TYPE_CHECKING, cast

from outlook_connector import config
from outlook_connector.domain.errors import AuthenticationRequired, ConnectorError

if TYPE_CHECKING:
    from outlook_connector.auth.tokens import AccessToken, CacheStatus, TokenProvider

EXIT_OK = 0
EXIT_ERROR = 1
EXIT_AUTH_REQUIRED = 2
EXIT_INTERRUPTED = 130


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.unsecure:
        print(
            "WARNING: --unsecure stores reusable Microsoft tokens as PLAINTEXT at "
            f"{config.token_cache_path(unsecure=True)}. Do not share it; delete it when done.",
            file=sys.stderr,
        )
    match args.command:
        case "auth":
            return _auth(args)
        case "status":
            return _status(args)
        case "mcp":
            return _mcp(args)
        case "ui":
            return _ui(args)
        case _:  # argparse accepts only the commands above
            raise AssertionError(f"unknown command {args.command!r}")


# ---------------------------------------------------------------------- handlers

# The decorator comes first: it runs when the handlers below are defined.
Handler = Callable[[argparse.Namespace], int]


def _exit_codes(handler: Handler) -> Handler:
    """For the terminal commands (auth, status): print an error and turn it into an exit code.
    The MCP server and the UI report their own errors (the MCP server's stdout is the protocol)."""

    @functools.wraps(handler)
    def wrapper(args: argparse.Namespace) -> int:
        try:
            return handler(args)
        except AuthenticationRequired as exc:
            print(str(exc), file=sys.stderr)
            return EXIT_AUTH_REQUIRED
        except ConnectorError as exc:
            print(f"Error: {exc}", file=sys.stderr)
            return EXIT_ERROR
        except KeyboardInterrupt:
            print("Interrupted.", file=sys.stderr)
            return EXIT_INTERRUPTED

    return wrapper


def _mcp(args: argparse.Namespace) -> int:
    """``mcp``: the MCP server over stdio; stdout belongs to the protocol from here on."""
    from outlook_connector.surfaces.mcp_main import serve_mcp

    serve_mcp(unsecure=args.unsecure)
    return EXIT_OK


def _ui(args: argparse.Namespace) -> int:
    """``ui``: the local web UI until Ctrl+C or idle."""
    from outlook_connector.surfaces.web.web_main import serve_ui

    with contextlib.suppress(KeyboardInterrupt):
        serve_ui(
            unsecure=args.unsecure,
            port=args.port,
            open_browser=not args.no_browser,
            idle_minutes=args.idle_minutes,
        )
    return EXIT_OK


@_exit_codes
def _auth(args: argparse.Namespace) -> int:
    """``auth``: device-code sign-in for one profile, or sign out."""
    provider = _tokens(args)
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


@_exit_codes
def _status(args: argparse.Namespace) -> int:
    """``status``: the signed-in account and profiles; ``--check`` also gets each token silently."""
    provider = _tokens(args)
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
    # Only the read profile is required; the write profile is optional until send/mutations exist.
    read_ok = checks.get("read", {"ok": True})["ok"]
    return EXIT_OK if read_ok else EXIT_AUTH_REQUIRED


def _tokens(args: argparse.Namespace) -> TokenProvider:
    from outlook_connector.auth.tokens import TokenProvider

    return TokenProvider(unsecure=args.unsecure)


# ---------------------------------------------------------------------- parser


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

    ui = sub.add_parser("ui", help="Start the local web UI (stops when idle).")
    ui.add_argument(
        "--port", type=int, default=8765, help="Preferred port (default 8765; a free one if busy)."
    )
    ui.add_argument("--no-browser", action="store_true", help="Do not open a browser window.")
    ui.add_argument("--idle-minutes", type=float, default=30, help="Stop after this many idle minutes.")
    _add_unsecure(ui)
    return parser


def _add_unsecure(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--unsecure", action="store_true", help="Development only: use a separate PLAINTEXT token cache."
    )


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
        f"  mail scopes: {', '.join(cast(list[str], facts['mail_scopes'])) or '(none)'}  "
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
                    f"mail scopes {', '.join(cast(list[str], check['mail_scopes'])) or '(none)'}"
                )
            else:
                print(f"  check: FAILED - {check['error']}")


def _iso(epoch: int | None) -> str | None:
    return datetime.fromtimestamp(epoch, UTC).isoformat(timespec="seconds") if epoch else None


if __name__ == "__main__":
    raise SystemExit(main())
