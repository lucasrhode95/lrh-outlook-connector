from __future__ import annotations

import asyncio
import time
from threading import Event

import pytest

from outlook_connector.surfaces.web import web_main


@pytest.fixture(autouse=True)
def keep_process_signals(monkeypatch) -> None:
    """Shutdown ignores Ctrl+C for the rest of the process; the test process keeps its own handler."""
    monkeypatch.setattr(web_main.signal, "signal", lambda *_args: None)


def test_ui_opens_browser_after_server_is_ready(monkeypatch, capsys) -> None:
    instances = []
    opened: list[str] = []
    browser_called = Event()

    class Server:
        def __init__(self, _config) -> None:
            self.started = False
            self.should_exit = False
            instances.append(self)

        async def serve(self) -> None:
            self.started = True
            while not self.should_exit:
                await asyncio.sleep(0)

    monkeypatch.setattr(web_main, "_free_port", lambda _port: 45678)
    monkeypatch.setattr(web_main, "AppContext", lambda **_kwargs: _Context())
    monkeypatch.setattr(web_main, "create_app", lambda *_args, **_kwargs: object())
    monkeypatch.setattr(web_main.uvicorn, "Server", Server)

    def open_browser(url: str) -> bool:
        assert instances[0].started
        opened.append(url)
        browser_called.set()
        instances[0].should_exit = True
        return True

    monkeypatch.setattr(web_main.webbrowser, "open", open_browser)
    web_main.serve_ui(unsecure=False, open_browser=True)

    assert browser_called.wait(timeout=5)
    assert opened == ["http://127.0.0.1:45678/"]
    output = capsys.readouterr().out
    assert "http://127.0.0.1:45678/" in output
    assert "did not start" not in output


def test_ui_reports_chosen_address_when_startup_fails(monkeypatch, capsys) -> None:
    opened: list[str] = []

    class Server:
        def __init__(self, _config) -> None:
            self.started = False
            self.should_exit = False

        async def serve(self) -> None:
            self.should_exit = True

    def record_browser(url: str) -> bool:
        opened.append(url)
        return True

    monkeypatch.setattr(web_main, "_free_port", lambda _port: 45679)
    monkeypatch.setattr(web_main, "AppContext", lambda **_kwargs: _Context())
    monkeypatch.setattr(web_main, "create_app", lambda *_args, **_kwargs: object())
    monkeypatch.setattr(web_main.uvicorn, "Server", Server)
    monkeypatch.setattr(web_main.webbrowser, "open", record_browser)

    web_main.serve_ui(unsecure=False, open_browser=True)

    output = capsys.readouterr().out
    assert "http://127.0.0.1:45679/" in output
    assert "did not start" in output
    assert opened == []


def test_a_browser_launcher_that_waits_for_the_browser_does_not_hold_up_the_ui(monkeypatch) -> None:
    launcher_started = Event()
    browser_closed = Event()

    class Server:
        def __init__(self, _config) -> None:
            self.started = False
            self.should_exit = False

        async def serve(self) -> None:
            self.started = True
            # Stop once the launcher is waiting on the browser: the UI must still shut down.
            while not launcher_started.is_set():
                await asyncio.sleep(0.01)
            self.should_exit = True

    def waiting_launcher(_url: str) -> bool:
        launcher_started.set()
        browser_closed.wait(timeout=30)  # stands in for a launcher that waits for the browser to close
        return True

    monkeypatch.setattr(web_main, "_free_port", lambda _port: 45680)
    monkeypatch.setattr(web_main, "AppContext", lambda **_kwargs: _Context())
    monkeypatch.setattr(web_main, "create_app", lambda *_args, **_kwargs: object())
    monkeypatch.setattr(web_main.uvicorn, "Server", Server)
    monkeypatch.setattr(web_main.webbrowser, "open", waiting_launcher)

    started = time.monotonic()
    web_main.serve_ui(unsecure=False, open_browser=True)

    assert time.monotonic() - started < 10
    assert not browser_closed.is_set()
    browser_closed.set()


class _Context:
    async def aclose(self) -> None:
        return None
