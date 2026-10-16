from __future__ import annotations

import io
import logging
from collections.abc import Iterator
from pathlib import Path

import pytest

from outlook_connector.logs import SecretRedactionFilter, configure_app_logging


@pytest.fixture
def root_logging() -> Iterator[None]:
    """Restore the root logger (and the quieted libraries) that configure_app_logging replaces."""
    root = logging.getLogger()
    handlers, level = root.handlers[:], root.level
    library_levels = {name: logging.getLogger(name).level for name in ("httpx", "httpcore")}
    yield
    for handler in root.handlers[:]:
        root.removeHandler(handler)
        handler.close()
    for handler in handlers:
        root.addHandler(handler)
    root.setLevel(level)
    for name, library_level in library_levels.items():
        logging.getLogger(name).setLevel(library_level)
    logging.captureWarnings(False)


def test_redaction_removes_synthetic_credentials() -> None:
    record = logging.LogRecord(
        "synthetic",
        logging.ERROR,
        "test_logs.py",
        1,
        "Authorization: Bearer token-value access_token=secret-value https://example.test/?code=query-value",
        (),
        None,
    )

    assert SecretRedactionFilter().filter(record)
    message = record.getMessage()
    assert "[REDACTED]" in message
    assert all(secret not in message for secret in ("token-value", "secret-value", "query-value"))


@pytest.mark.usefixtures("root_logging")
def test_mcp_logs_to_its_file_only(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    path = configure_app_logging("mcp", directory=tmp_path)
    logging.getLogger("outlook_connector.synthetic").info("synthetic event")

    assert path.parent == tmp_path and path.name.startswith("mcp-")
    assert "synthetic event" in path.read_text(encoding="utf-8")
    assert capsys.readouterr().out == ""


@pytest.mark.usefixtures("root_logging")
def test_ui_also_logs_to_the_stream(tmp_path: Path) -> None:
    stream = io.StringIO()
    path = configure_app_logging("ui", stdout=True, stream=stream, directory=tmp_path)
    logging.getLogger("outlook_connector.synthetic").warning("synthetic warning Bearer abc.def")

    assert "synthetic warning Bearer [REDACTED]" in stream.getvalue()
    assert "synthetic warning Bearer [REDACTED]" in path.read_text(encoding="utf-8")


@pytest.mark.usefixtures("root_logging")
def test_request_urls_from_http_libraries_stay_out(tmp_path: Path) -> None:
    path = configure_app_logging("mcp", directory=tmp_path)
    logging.getLogger("httpx").info("HTTP Request: GET https://example.test/me/messages?$search=synthetic")

    assert "synthetic" not in path.read_text(encoding="utf-8")


def test_credential_headers_are_redacted_in_full() -> None:
    record = logging.LogRecord(
        "synthetic",
        logging.INFO,
        "test_logs.py",
        1,
        "Cookie: sid=one-value; refresh=two-value\nAuthorization: Basic three-value",
        (),
        None,
    )

    assert SecretRedactionFilter().filter(record)
    message = record.getMessage()
    assert message == "Cookie: [REDACTED]\nAuthorization: [REDACTED]"


def test_component_must_be_a_short_name(tmp_path: Path) -> None:
    with pytest.raises(ValueError):
        configure_app_logging("../escape", directory=tmp_path)
