"""`outlook-connector ui`: start the local web UI on demand and exit when idle.

Imported only by the `ui` command, so MCP processes never load the web stack.
"""

from __future__ import annotations

import asyncio
import contextlib
import secrets
import signal
import socket
import threading
import webbrowser

import uvicorn

from outlook_connector.bootstrap import AppContext
from outlook_connector.surfaces.ui_settings import DEFAULT_UI_IDLE_MINUTES, DEFAULT_UI_PORT
from outlook_connector.surfaces.web.routes import Activity, create_app


def serve_ui(
    *,
    unsecure: bool,
    port: int = DEFAULT_UI_PORT,
    open_browser: bool = True,
    idle_minutes: float = DEFAULT_UI_IDLE_MINUTES,
) -> None:
    port = _free_port(port)
    token = secrets.token_urlsafe(24)
    activity = Activity()
    context = AppContext(unsecure=unsecure)
    app = create_app(context, session_token=token, port=port, activity=activity)
    # lifespan off: the app has no startup/shutdown hooks, and Ctrl+C would log the cancelled task as an error
    config = uvicorn.Config(app, host="127.0.0.1", port=port, log_level="warning", lifespan="off")
    server = uvicorn.Server(config)
    url = f"http://127.0.0.1:{port}/"

    async def serve() -> None:
        async def watch_idle() -> None:
            while not server.should_exit:
                await asyncio.sleep(15)
                if activity.idle_seconds() > idle_minutes * 60:
                    print(f"No activity for {idle_minutes:g} minutes; stopping.", flush=True)
                    server.should_exit = True

        watcher = asyncio.create_task(watch_idle())
        server_task = asyncio.create_task(server.serve())
        try:
            # The browser opens once the server answers, so a failed start is reported instead.
            while not server.started and not server_task.done():
                await asyncio.sleep(0.05)
            if not server.started:
                print(
                    f"Outlook connector UI did not start at {url}; check the startup error above.",
                    flush=True,
                )
            elif open_browser:
                _open_browser(url)
            await server_task
        finally:
            watcher.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await watcher
            if not server_task.done():
                server.should_exit = True
                await server_task
            await _close(context)

    print(
        f"Outlook connector UI: {url}  (Ctrl+C to stop; stops after {idle_minutes:g} idle minutes)",
        flush=True,
    )
    asyncio.run(serve())


def _open_browser(url: str) -> None:
    """Open the browser on a daemon thread. A launcher that waits for the browser must not hold up the UI:
    the thread is never awaited, and it does not keep the process alive."""

    def open_or_report() -> None:
        if not webbrowser.open(url):
            print(f"Open this address in your browser: {url}", flush=True)

    threading.Thread(target=open_or_report, daemon=True).start()


def _free_port(preferred: int) -> int:
    for port in (preferred, 0):
        with contextlib.closing(socket.socket(socket.AF_INET, socket.SOCK_STREAM)) as probe:
            try:
                probe.bind(("127.0.0.1", port))
            except OSError:
                continue
            return probe.getsockname()[1]
    raise OSError("No free local port.")


async def _close(context: AppContext) -> None:
    """Close the connections even when Ctrl+C interrupts the shutdown. Further presses are ignored
    from here (closing sockets takes well under a second), and the close runs as its own task,
    shielded from the cancellation an earlier press already requested, and is waited for."""
    if threading.current_thread() is threading.main_thread():
        signal.signal(signal.SIGINT, signal.SIG_IGN)
    closing = asyncio.ensure_future(context.aclose())
    while not closing.done():
        with contextlib.suppress(asyncio.CancelledError):
            await asyncio.shield(closing)
