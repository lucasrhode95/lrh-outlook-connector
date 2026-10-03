from __future__ import annotations

from pathlib import Path

import pytest

from outlook_connector import config


def test_profiles_never_request_a_denied_pair() -> None:
    for profile in config.PROFILES.values():
        for scope in profile.scopes:
            assert (profile.client_id, scope) not in config.DENIED_PAIRS


def test_read_and_write_use_the_researched_clients() -> None:
    assert config.PROFILES["read"].client_id == config.OUTLOOK_MOBILE_CLIENT_ID
    assert config.PROFILES["read"].scopes == ("https://graph.microsoft.com/Mail.Read",)
    assert config.PROFILES["write"].client_id == config.ONE_OUTLOOK_WEB_CLIENT_ID
    assert config.PROFILES["write"].scopes == ("https://outlook.office.com/.default",)


def test_cache_paths_live_in_the_data_dir_and_differ_by_mode(isolated_home: Path) -> None:
    secure = config.token_cache_path(unsecure=False)
    plain = config.token_cache_path(unsecure=True)
    assert secure.parent == plain.parent == isolated_home.resolve()
    assert secure != plain
    assert "plaintext" in plain.name


def test_windows_data_dir_is_outside_appdata(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    # Packaged apps (the Claude desktop app and the MCP servers it starts) get AppData redirected.
    monkeypatch.delenv("OUTLOOK_CONNECTOR_HOME")
    monkeypatch.setattr(config, "_windows", lambda: True)
    monkeypatch.setattr(Path, "home", classmethod(lambda cls: tmp_path / "rhode"))
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "rhode" / "AppData" / "Local"))
    assert config.data_dir() == tmp_path / "rhode" / ".lrh-outlook-connector"
