from __future__ import annotations

import json

import pytest

from outlook_connector import __main__ as cli
from outlook_connector import config
from outlook_connector.auth.tokens import TokenProvider
from tests.fakes.msal_fakes import Script, msal_account, seed_cache, token_result, use_script

GRAPH = "https://graph.microsoft.com"
GRAPH_CLIENT = config.OUTLOOK_MOBILE_CLIENT_ID


def run(argv: list[str], script: Script) -> int:
    use_script(script)  # the CLI's sign-ins go to the scripted fake MSAL (conftest)
    return cli.main(argv)


def test_ui_defaults_use_shared_values() -> None:
    from outlook_connector.surfaces.ui_settings import DEFAULT_UI_IDLE_MINUTES, DEFAULT_UI_PORT

    args = cli.build_parser().parse_args(["ui"])
    assert args.port == DEFAULT_UI_PORT
    assert args.idle_minutes == DEFAULT_UI_IDLE_MINUTES


def test_auth_signs_in_with_device_code(capsys: pytest.CaptureFixture[str]) -> None:
    script = Script(
        device_result=token_result(
            aud=GRAPH, scp="Mail.Read Mail.Read.Shared User.Read", source="identity_provider"
        )
    )
    assert run(["auth", "graph", "--unsecure"], script) == cli.EXIT_OK
    out = capsys.readouterr()
    assert "ABC123" in out.out
    assert "Mail.Read, Mail.Read.Shared" in out.out
    assert "PLAINTEXT" in out.err  # the --unsecure warning is always printed


def test_auth_skips_device_code_when_already_signed_in(capsys: pytest.CaptureFixture[str]) -> None:
    script = Script(
        accounts=[msal_account()], silent={GRAPH_CLIENT: token_result(aud=GRAPH, scp="Mail.Read")}
    )
    assert run(["auth", "--unsecure"], script) == cli.EXIT_OK
    assert "Already signed in" in capsys.readouterr().out
    assert script.shown_flows == 0


def test_auth_force_always_runs_device_code() -> None:
    script = Script(
        accounts=[msal_account()],
        silent={GRAPH_CLIENT: token_result(aud=GRAPH, scp="Mail.Read")},
        device_result=token_result(aud=GRAPH, scp="Mail.Read"),
    )
    assert run(["auth", "graph", "--force", "--unsecure"], script) == cli.EXIT_OK
    assert script.shown_flows == 1


def test_status_json_is_offline_and_parseable(capsys: pytest.CaptureFixture[str]) -> None:
    provider = TokenProvider(unsecure=True)
    seed_cache(provider._cache(), client_id=GRAPH_CLIENT, scope="https://graph.microsoft.com/Mail.Read")
    script = Script()
    assert run(["status", "--json", "--unsecure"], script) == cli.EXIT_OK
    data = json.loads(capsys.readouterr().out)
    assert [p["profile"] for p in data["profiles"]] == ["graph", "outlook"]
    assert data["cache"]["mode"] == "plaintext-development"
    assert [p["signed_in"] for p in data["profiles"]] == [True, False]
    assert script.created == []  # no MSAL app (and so no network) for offline status


def test_status_check_reports_failures_but_only_graph_is_required(capsys: pytest.CaptureFixture[str]) -> None:
    script = Script(
        accounts=[msal_account()], silent={GRAPH_CLIENT: token_result(aud=GRAPH, scp="Mail.Read")}
    )
    assert run(["status", "--check", "--unsecure"], script) == cli.EXIT_OK
    out = capsys.readouterr().out
    assert "check: ok" in out
    assert "check: FAILED" in out  # the Outlook profile has no token


def test_missing_sign_in_exit_code_and_message(capsys: pytest.CaptureFixture[str]) -> None:
    script = Script(device_result={"error": "authorization_declined"})
    assert run(["auth", "outlook", "--unsecure"], script) == cli.EXIT_ERROR
    assert "authorization_declined" in capsys.readouterr().err


def test_sign_out(capsys: pytest.CaptureFixture[str]) -> None:
    assert run(["auth", "--sign-out", "--unsecure"], Script()) == cli.EXIT_OK
    assert "No token cache" in capsys.readouterr().out


@pytest.mark.parametrize("profile", ["graph", "outlook"])
def test_auth_parser_accepts_service_profiles(profile: str) -> None:
    assert cli.build_parser().parse_args(["auth", profile]).profile == profile


@pytest.mark.parametrize("profile", ["read", "write"])
def test_auth_parser_rejects_old_profile_names(profile: str) -> None:
    with pytest.raises(SystemExit) as error:
        cli.build_parser().parse_args(["auth", profile])
    assert error.value.code == 2
