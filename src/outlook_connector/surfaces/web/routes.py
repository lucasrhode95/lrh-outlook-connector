"""Local web UI: a Starlette JSON API over the shared service, plus the static page.

Security: bound to 127.0.0.1 by the launcher; every /api request must carry the per-run session
token and a localhost Host header (blocks other local pages and DNS rebinding). Mail content is
untrusted; the page renders it as text only.
"""

from __future__ import annotations

import time
from collections.abc import Awaitable, Callable
from datetime import datetime
from importlib import resources
from typing import Any
from urllib.parse import quote

from pydantic import ValidationError
from starlette.applications import Starlette
from starlette.middleware import Middleware
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request
from starlette.responses import FileResponse, HTMLResponse, JSONResponse, Response
from starlette.routing import Mount, Route
from starlette.staticfiles import StaticFiles

from outlook_connector.bootstrap import AppContext
from outlook_connector.domain.errors import (
    AuthenticationRequired,
    ConnectorError,
    InvalidRequest,
    NotFound,
    Throttled,
)
from outlook_connector.domain.models import ExportRequest, Scope

STATIC = resources.files("outlook_connector.surfaces.web") / "static"
_STATUS = {AuthenticationRequired: 401, InvalidRequest: 400, NotFound: 404, Throttled: 429}


class Activity:
    def __init__(self) -> None:
        self.last = time.monotonic()

    def touch(self) -> None:
        self.last = time.monotonic()

    def idle_seconds(self) -> float:
        return time.monotonic() - self.last


def _guard(token: str, port: int, activity: Activity) -> type[BaseHTTPMiddleware]:
    allowed_hosts = {f"127.0.0.1:{port}", f"localhost:{port}"}

    class Guard(BaseHTTPMiddleware):
        async def dispatch(
            self, request: Request, call_next: Callable[[Request], Awaitable[Response]]
        ) -> Response:
            if request.headers.get("host") not in allowed_hosts:
                return JSONResponse({"error": "Forbidden host."}, status_code=403)
            if request.url.path.startswith("/api/") and request.headers.get("x-session-token") != token:
                return JSONResponse({"error": "Missing or wrong session token."}, status_code=403)
            activity.touch()
            return await call_next(request)

    return Guard


def _flag(request: Request, name: str, default: bool = False) -> bool:
    value = request.query_params.get(name)
    return default if value is None else value.lower() in ("1", "true", "yes")


def _scope(request: Request) -> Scope:
    """Validate the web query's scope once at the route entry point."""
    return Scope(
        sent_items=_flag(request, "sent_items", True),
        meeting_mail=_flag(request, "meeting_mail", True),
        deleted_items=_flag(request, "deleted_items"),
    )


def _when(request: Request, name: str) -> datetime | None:
    value = request.query_params.get(name)
    if not value:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        raise InvalidRequest(f"{name} must be an ISO 8601 date-time.") from None


def _int(request: Request, name: str, default: int) -> int:
    try:
        return int(request.query_params.get(name, default))
    except ValueError:
        raise InvalidRequest(f"{name} must be an integer.") from None


def _json(model: Any) -> JSONResponse:
    if isinstance(model, list):
        return JSONResponse([m.model_dump(mode="json") for m in model])
    return JSONResponse(model.model_dump(mode="json"))


def create_app(context: AppContext, *, session_token: str, port: int, activity: Activity) -> Starlette:
    index_html = (
        (STATIC / "index.html").read_text(encoding="utf-8").replace("{{SESSION_TOKEN}}", session_token)
    )

    def api(handler: Callable[[Request], Awaitable[Response]]) -> Callable[[Request], Awaitable[Response]]:
        async def wrapped(request: Request) -> Response:
            try:
                return await handler(request)
            except ConnectorError as exc:
                status = next((code for kind, code in _STATUS.items() if isinstance(exc, kind)), 502)
                return JSONResponse({"error": str(exc), "kind": type(exc).__name__}, status_code=status)
            except ValidationError as exc:
                return JSONResponse({"error": str(exc), "kind": "InvalidRequest"}, status_code=400)

        return wrapped

    async def index(_: Request) -> Response:
        return HTMLResponse(index_html, headers={"Cache-Control": "no-store"})

    async def status(_: Request) -> Response:
        cache = context.tokens.status()
        return JSONResponse(
            {
                "account": next((a.username for a in cache.accounts), None),
                "signed_in": {p.profile: p.signed_in for p in cache.profiles},
                "sign_in_command": context.tokens.sign_in_command("graph"),
            }
        )

    async def me(_: Request) -> Response:
        return _json(await (await context.services()).mailbox.profile())

    async def photo(_: Request) -> Response:
        data = await (await context.services()).mailbox.photo()
        if data is None:
            return JSONResponse({"error": "No profile photo.", "kind": "NotFound"}, status_code=404)
        return Response(data, media_type="image/jpeg", headers={"Cache-Control": "private, max-age=3600"})

    async def folders(request: Request) -> Response:
        return _json(await (await context.services()).mailbox.folders(refresh=_flag(request, "refresh")))

    async def messages(request: Request) -> Response:
        page = await (await context.services()).mailbox.list_messages(
            folder=request.query_params.get("folder") or None,
            since=_when(request, "since"),
            until=_when(request, "until"),
            limit=_int(request, "limit", 100),
            cursor=request.query_params.get("cursor") or None,
            scope=_scope(request),
        )
        return _json(page)

    async def search(request: Request) -> Response:
        result = await (await context.services()).mailbox.search(
            request.query_params.get("q", ""),
            since=_when(request, "since"),
            until=_when(request, "until"),
            folder=request.query_params.get("folder") or None,
            limit=_int(request, "limit", 50),
            cursor=request.query_params.get("cursor") or None,
            scope=_scope(request),
        )
        return _json(result)

    async def conversation(request: Request) -> Response:
        result = await (await context.services()).conversations.get_conversation(
            request.path_params["conversation_id"],
            include_bodies=False,
            scope=_scope(request),
        )
        return _json(result)

    async def conversation_sizes(request: Request) -> Response:
        payload = await request.json()
        ids = payload.get("conversation_ids") if isinstance(payload, dict) else None
        if not isinstance(ids, list) or not all(isinstance(i, str) for i in ids):
            raise InvalidRequest("conversation_ids must be a list of strings.")
        sizes = await (await context.services()).mailbox.conversation_sizes(
            ids, scope=Scope.model_validate(payload.get("scope", {}))
        )
        return _json(sizes)

    async def message(request: Request) -> Response:
        body = request.query_params.get("body", "unique")
        if body not in ("unique", "full"):
            raise InvalidRequest("body must be unique or full.")
        content = await (await context.services()).mailbox.get_message(
            request.path_params["message_id"],
            body=body,  # type: ignore[arg-type]
            max_chars=None,  # the whole body in one response: fetched from the server once
        )
        return _json(content)

    async def attachment(request: Request) -> Response:
        saved = await (await context.services()).files.download_attachment(
            request.path_params["message_id"], request.path_params["attachment_id"]
        )
        return FileResponse(saved.path, media_type=saved.content_type, filename=saved.name)

    async def export(request: Request) -> Response:
        export_request = ExportRequest.model_validate(await request.json())
        artifact = await (await context.services()).exports.export(export_request)
        # The file's "Export errors: ..." line, for the UI to show after the download.
        headers = {"X-Export-Errors": quote(artifact.error_summary)} if artifact.error_summary else None
        return FileResponse(
            artifact.path, media_type=artifact.content_type, filename=artifact.filename, headers=headers
        )

    async def attachment_names(request: Request) -> Response:
        payload = await request.json()
        ids = [str(i) for i in payload.get("message_ids") or []]
        found = await (await context.services()).mailbox.attachments_many(ids)
        return JSONResponse(
            {
                mid: [a.model_dump(mode="json") for a in items if not a.is_inline]
                for mid, items in found.items()
            }
        )

    async def heartbeat(_: Request) -> Response:
        return JSONResponse({"ok": True})

    routes = [
        Route("/", index),
        Route("/api/status", api(status)),
        Route("/api/me", api(me)),
        Route("/api/me/photo", api(photo)),
        Route("/api/folders", api(folders)),
        Route("/api/messages", api(messages)),
        Route("/api/messages/{message_id}", api(message)),
        Route("/api/messages/{message_id}/attachments/{attachment_id}", api(attachment)),
        Route("/api/search", api(search)),
        Route("/api/conversation-sizes", api(conversation_sizes), methods=["POST"]),
        Route("/api/attachments", api(attachment_names), methods=["POST"]),
        Route("/api/conversations/{conversation_id}", api(conversation)),
        Route("/api/export", api(export), methods=["POST"]),
        Route("/api/heartbeat", heartbeat, methods=["POST"]),
        Mount("/static", StaticFiles(directory=str(STATIC)), name="static"),
    ]
    return Starlette(routes=routes, middleware=[Middleware(_guard(session_token, port, activity))])
