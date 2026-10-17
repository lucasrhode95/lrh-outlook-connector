"""Command-line entry point: ``outlook-connector <command>``.

Top-down: ``main`` first, then one handler per command, then the parser and the output helpers.
Handlers import their dependencies lazily, so that ``mcp`` never loads the web stack.
"""

from __future__ import annotations

import argparse
import contextlib
import functools
import sys
from collections.abc import Callable
from datetime import UTC, datetime
from typing import cast

from m365_auth import AccessToken, TokenProvider, check_profiles, status_lines

from outlook_connector import config
from outlook_connector.domain.errors import AuthenticationRequired, ConnectorError
from outlook_connector.surfaces.ui_settings import DEFAULT_UI_IDLE_MINUTES, DEFAULT_UI_PORT

EXIT_OK = 0
EXIT_ERROR = 1
EXIT_AUTH_REQUIRED = 2
EXIT_INTERRUPTED = 130


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.unsecure:
        print(
            "WARNING: --unsecure stores reusable Microsoft tokens in a separate PLAINTEXT shared-auth cache. "
            "Do not share it; delete it when done.",
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
    from outlook_connector.logs import configure_app_logging
    from outlook_connector.surfaces.mcp_main import serve_mcp

    configure_app_logging("mcp")  # file only
    serve_mcp(unsecure=args.unsecure)
    return EXIT_OK


def _ui(args: argparse.Namespace) -> int:
    """``ui``: the local web UI until Ctrl+C or idle."""
    from outlook_connector.logs import configure_app_logging
    from outlook_connector.surfaces.web.web_main import serve_ui

    configure_app_logging("ui", stdout=True)
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
    """``auth``: prepare every shared Microsoft profile, or sign out of all connectors."""
    provider = _tokens(args)
    if args.sign_out:
        existed = provider.sign_out()
        print("Signed out: token cache deleted." if existed else "No token cache to delete.")
        return EXIT_OK
    results = provider.bootstrap(show=lambda message: print(f"\n{message}\n", flush=True))
    for result in results:
        purpose = provider.profiles[result.profile].purpose
        if result.ready and result.token is not None:
            print(f"[{result.profile}] Ready: {purpose}.")
            _print_token_summary(result.token)
        else:
            print(f"[{result.profile}] Not ready: {purpose}. {result.error}", file=sys.stderr)
    if all(result.ready for result in results):
        return EXIT_OK
    first_error = next((result.error for result in results if result.error is not None), None)
    return EXIT_AUTH_REQUIRED if isinstance(first_error, AuthenticationRequired) else EXIT_ERROR


@_exit_codes
def _status(args: argparse.Namespace) -> int:
    """Show cached client credentials; ``--check`` also attempts each profile's token silently."""
    provider = _tokens(args)
    checks = check_profiles(provider) if args.check else {}
    print("\n".join(status_lines(provider, checks)))
    # Core mail data needs Graph; inbox rules, signatures and mail changes need Outlook.
    graph_ok = checks.get("outlook.graph", {"ok": True})["ok"]
    return EXIT_OK if graph_ok else EXIT_AUTH_REQUIRED


def _tokens(args: argparse.Namespace) -> TokenProvider:
    return TokenProvider(unsecure=args.unsecure, login_command=f"{config.CLI_NAME} auth")


# ---------------------------------------------------------------------- parser


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog=config.CLI_NAME, description="Local Outlook mailbox connector.")
    sub = parser.add_subparsers(dest="command", required=True)

    auth = sub.add_parser(
        "auth", help="Prepare Microsoft 365 profiles with device codes (the only command that prompts)."
    )
    auth.add_argument("--sign-out", action="store_true", help="Sign out of the shared Microsoft account.")
    _add_unsecure(auth)

    status = sub.add_parser(
        "status", help="Show cached client credentials and resource-token state (offline by default)."
    )
    status.add_argument(
        "--check",
        action="store_true",
        help="Also acquire each profile's token silently; this checks token acquisition, not API access.",
    )
    _add_unsecure(status)

    mcp = sub.add_parser("mcp", help="Run the MCP server over stdio (started by your MCP client).")
    _add_unsecure(mcp)

    ui = sub.add_parser("ui", help="Start the local web UI (stops when idle).")
    ui.add_argument(
        "--port",
        type=int,
        default=DEFAULT_UI_PORT,
        help=f"Preferred port (default {DEFAULT_UI_PORT}; a free one if busy).",
    )
    ui.add_argument("--no-browser", action="store_true", help="Do not open a browser window.")
    ui.add_argument(
        "--idle-minutes",
        type=float,
        default=DEFAULT_UI_IDLE_MINUTES,
        help=f"Stop after this many idle minutes (default {DEFAULT_UI_IDLE_MINUTES}).",
    )
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
    mail_scopes = cast(list[str], facts["mail_scopes"])
    if mail_scopes:
        print(f"  mail scopes: {', '.join(mail_scopes)}")
    print(f"  granted scope count: {facts['scope_count']}")
    print(f"  token from:  {facts['source']}, expires {facts['expires_on']}")


def _iso(epoch: int | None) -> str | None:
    return datetime.fromtimestamp(epoch, UTC).isoformat(timespec="seconds") if epoch else None


if __name__ == "__main__":
    raise SystemExit(main())
