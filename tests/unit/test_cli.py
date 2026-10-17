from __future__ import annotations

import pytest
from m365_auth import PROFILES, TokenProvider

from outlook_connector import __main__ as cli
from tests.fakes.msal_fakes import Script, msal_account, seed_cache, token_result, use_script

GRAPH = "https://graph.microsoft.com"
GRAPH_CLIENT = PROFILES["outlook.graph"].client_id


def run(argv: list[str], script: Script) -> int:
    use_script(script)
    return cli.main(argv)


def test_ui_defaults_use_shared_values() -> None:
    from outlook_connector.surfaces.ui_settings import DEFAULT_UI_IDLE_MINUTES, DEFAULT_UI_PORT

    args = cli.build_parser().parse_args(["ui"])
    assert args.port == DEFAULT_UI_PORT
    assert args.idle_minutes == DEFAULT_UI_IDLE_MINUTES


def test_auth_bootstraps_all_products_with_profile_appropriate_scopes(
    capsys: pytest.CaptureFixture[str],
) -> None:
    graph = token_result(aud=GRAPH, scp="Mail.Read", source="identity_provider")
    outlook = token_result(aud="https://outlook.office.com", scp="Mail.ReadWrite", source="identity_provider")
    search = token_result(
        aud="https://outlook.office.com", scp="SubstrateSearch-Internal.ReadWrite", source="identity_provider"
    )
    spaces = token_result(aud="https://api.spaces.skype.com", scp="chat", source="identity_provider")
    csa = token_result(aud="https://chatsvcagg.teams.microsoft.com", scp="chat", source="cache")
    teams_graph = token_result(aud=GRAPH, scp="User.Read", source="cache")
    script = Script(
        silent_by_scopes={
            PROFILES["teams.csa"].scopes: csa,
            PROFILES["teams.graph"].scopes: teams_graph,
        },
        device_by_scopes={
            PROFILES["outlook.graph"].scopes: graph,
            PROFILES["outlook.api"].scopes: outlook,
            PROFILES["m365.search"].scopes: search,
            PROFILES["teams.spaces"].scopes: spaces,
        },
    )
    assert run(["auth", "--unsecure"], script) == cli.EXIT_OK
    out = capsys.readouterr()
    assert "ABC123" in out.out
    assert "[outlook.graph] Ready" in out.out
    assert "[m365.search] Ready" in out.out
    assert "[teams.graph] Ready" in out.out
    assert "PLAINTEXT" in out.err
    assert script.shown_flows == 4  # uncached resources need separate device flows


def test_status_is_offline_and_lists_every_profile(capsys: pytest.CaptureFixture[str]) -> None:
    provider = TokenProvider(unsecure=True, login_command="outlook-connector auth")
    seed_cache(provider._cache(), client_id=GRAPH_CLIENT, scope="https://graph.microsoft.com/Mail.Read")
    script = Script()
    assert run(["status", "--unsecure"], script) == cli.EXIT_OK
    text = capsys.readouterr().out.splitlines()
    headings = [line.split(" ")[0] for line in text if line.startswith("[")]
    assert headings == [f"[{name}]" for name in PROFILES]
    assert "Token cache: plaintext-development" in " ".join(text)
    assert script.created == []


def test_status_check_reports_token_acquisition_without_claiming_api_access(
    capsys: pytest.CaptureFixture[str],
) -> None:
    script = Script(
        accounts=[msal_account()],
        silent={GRAPH_CLIENT: token_result(aud=GRAPH, scp="Mail.Read")},
    )
    assert run(["status", "--check", "--unsecure"], script) == cli.EXIT_OK
    out = capsys.readouterr().out
    assert "token acquisition ok" in out
    assert "not a service-permission test" in out
    assert "check: FAILED" in out


def test_declined_sign_in_stops_new_device_flows(capsys: pytest.CaptureFixture[str]) -> None:
    script = Script(device_result={"error": "authorization_declined"})
    assert run(["auth", "--unsecure"], script) == cli.EXIT_ERROR
    assert "authorization_declined" in capsys.readouterr().err
    assert script.shown_flows == 1


def test_sign_out(capsys: pytest.CaptureFixture[str]) -> None:
    assert run(["auth", "--sign-out", "--unsecure"], Script()) == cli.EXIT_OK
    assert "No token cache" in capsys.readouterr().out


@pytest.mark.parametrize("profile", ["graph", "outlook", "read", "write"])
def test_auth_parser_rejects_profile_arguments(profile: str) -> None:
    with pytest.raises(SystemExit) as error:
        cli.build_parser().parse_args(["auth", profile])
    assert error.value.code == 2


def test_status_parser_and_project_data_path_remain_available() -> None:
    with pytest.raises(SystemExit):  # the machine-readable form is the auth_status tool
        cli.build_parser().parse_args(["status", "--json"])
