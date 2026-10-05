from __future__ import annotations

from pathlib import Path

import msal
import pytest

from tests.fakes.msal_fakes import FAKE_MSAL, Script


@pytest.fixture(autouse=True)
def isolated_home(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    """Every test gets its own data directory; nothing touches the real token cache."""
    home = tmp_path / "home"
    monkeypatch.setenv("OUTLOOK_CONNECTOR_HOME", str(home))
    return home


@pytest.fixture(autouse=True)
def fake_msal(monkeypatch: pytest.MonkeyPatch) -> None:
    """Every test signs in against the scripted fake MSAL, never against Microsoft."""
    FAKE_MSAL.script = Script()
    monkeypatch.setattr(msal, "PublicClientApplication", FAKE_MSAL)
