"""`outlook-connector ui`: start the local web UI on demand and exit when idle.

Imported only by the `ui` command, so MCP processes never load the web stack.
"""

from __future__ import annotations

import asyncio
import contextlib
import secrets
import socket
import threading
import webbrowser

import uvicorn

from outlook_connector.bootstrap import AppContext
from outlook_connector.surfaces.web.routes import Activity, create_app

DEFAULT_PORT = 8765


def _free_port(preferred: int) -> int:
    for port in (preferred, 0):
        with contextlib.closing(socket.socket(socket.AF_INET, socket.SOCK_STREAM)) as probe:
            try:
                probe.bind(("127.0.0.1", port))
            except OSError:
                continue
            return probe.getsockname()[1]
    raise OSError("No free local port.")


def serve_ui(
    *, unsecure: bool, port: int = DEFAULT_PORT, open_browser: bool = True, idle_minutes: float = 30
) -> None:
    port = _free_port(port)
    token = secrets.token_urlsafe(24)
    activity = Activity()
    context = AppContext(unsecure=unsecure)
    app = create_app(context, session_token=token, port=port, activity=activity)
    server = uvicorn.Server(uvicorn.Config(app, host="127.0.0.1", port=port, log_level="warning"))
    url = f"http://127.0.0.1:{port}/"

    async def serve() -> None:
        async def watch_idle() -> None:
            while not server.should_exit:
                await asyncio.sleep(15)
                if activity.idle_seconds() > idle_minutes * 60:
                    print(f"No activity for {idle_minutes:g} minutes; stopping.", flush=True)
                    server.should_exit = True

        watcher = asyncio.create_task(watch_idle())
        try:
            await server.serve()
        finally:
            watcher.cancel()
            await context.aclose()

    print(
        f"Outlook connector UI: {url}  (Ctrl+C to stop; stops after {idle_minutes:g} idle minutes)",
        flush=True,
    )
    if open_browser:
        threading.Timer(1.0, webbrowser.open, args=(url,)).start()
    asyncio.run(serve())
