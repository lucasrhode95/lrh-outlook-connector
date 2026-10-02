from __future__ import annotations

from pathlib import Path

import pytest


@pytest.fixture(autouse=True)
def isolated_home(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    """Every test gets its own data directory; nothing touches the real token cache."""
    home = tmp_path / "home"
    monkeypatch.setenv("OUTLOOK_CONNECTOR_HOME", str(home))
    return home
