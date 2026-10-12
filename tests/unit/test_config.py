from __future__ import annotations

from pathlib import Path

import pytest
from m365_auth import PROFILES

from outlook_connector import config


def test_auth_profiles_have_product_qualified_names() -> None:
    assert PROFILES["outlook.graph"].scopes == (
        "https://graph.microsoft.com/Mail.Read",
        "https://graph.microsoft.com/User.Read",
    )
    assert PROFILES["outlook.api"].scopes == ("https://outlook.office.com/.default",)


def test_windows_data_dir_is_outside_appdata(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    # Packaged apps (the Claude desktop app and the MCP servers it starts) get AppData redirected.
    monkeypatch.delenv("OUTLOOK_CONNECTOR_HOME")
    monkeypatch.setattr(config, "_windows", lambda: True)
    monkeypatch.setattr(Path, "home", classmethod(lambda cls: tmp_path / "example-user"))
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "example-user" / "AppData" / "Local"))
    assert config.data_dir() == tmp_path / "example-user" / ".lrh-outlook-connector"
