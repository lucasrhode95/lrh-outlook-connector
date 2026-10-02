"""Shared plumbing for the research probes (stdlib only, no MSAL).

Authentication is a plain OAuth 2.0 device-code flow against Microsoft
first-party public clients, followed by refresh-token grants. Tokens are kept
as PLAINTEXT in ``.local/probe-tokens.json`` (git-ignored). This is a research
convenience only; the product will use an encrypted cache.

Output rules for every probe: print statuses, counts, shapes, error codes and
salted short hashes of identifiers. Never print tokens, bodies, subjects,
addresses, or raw identifiers.
"""

from __future__ import annotations

import base64
import hashlib
import json
import os
import secrets
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterator

ROOT = Path(__file__).resolve().parents[2]
LOCAL = ROOT / ".local"
TOKEN_FILE = LOCAL / "probe-tokens.json"
TENANT = "organizations"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0"
EXPECTED_USER = "lucas.rhode@landisgyr.com"

GRAPH = "https://graph.microsoft.com/v1.0"
OWS_URL = "https://outlook.cloud.microsoft/owa/service.svc"
IMMUTABLE = 'IdType="ImmutableId"'

READ_CLIENT = "27922004-5251-4030-b22d-91ecd9a37ea4"   # Outlook Mobile (Candidate C)
WRITE_CLIENT = "9199bf20-a13f-4107-85dc-02114787ef48"  # One Outlook Web (Candidate A)

# profile -> (client_id, resource scope)
PROFILES: dict[str, tuple[str, str]] = {
    "read": (READ_CLIENT, "https://graph.microsoft.com/Mail.Read"),
    "write": (WRITE_CLIENT, "https://outlook.office.com/.default"),
    "search": (WRITE_CLIENT, "https://outlook.office.com/search/.default"),
}
# Recorded AADSTS65002 denials. Never request these again.
DENIED = {
    ("d3590ed6-52b3-4102-aeff-aad2292ab01c", "https://graph.microsoft.com/Mail.Read"),
    (READ_CLIENT, "https://graph.microsoft.com/Mail.Send"),
    (READ_CLIENT, "https://graph.microsoft.com/Mail.ReadWrite"),          # 2026-10-02
    (READ_CLIENT, "https://graph.microsoft.com/Mail.ReadWrite.Shared"),   # 2026-10-02
    (WRITE_CLIENT, "https://graph.microsoft.com/Mail.Read"),              # 2026-10-02
    (WRITE_CLIENT, "https://graph.microsoft.com/Mail.ReadWrite"),         # 2026-10-02
    (WRITE_CLIENT, "https://graph.microsoft.com/Mail.Send"),              # 2026-10-02
}
ALLOWED_HOSTS = {
    "graph.microsoft.com",
    "outlook.office.com",
    "outlook.cloud.microsoft",
    "substrate.office.com",
    "login.microsoftonline.com",
}
_SALT = secrets.token_bytes(16)  # per-process: hashes correlate within one run only


# --------------------------------------------------------------------------- output

def emit(obj: Any) -> None:
    print(json.dumps(obj, indent=2, ensure_ascii=False, default=str), flush=True)


def tag(value: Any) -> str | None:
    """Short salted hash so identifiers can be correlated without being revealed."""
    if value is None:
        return None
    return hashlib.sha256(_SALT + str(value).encode("utf-8")).hexdigest()[:10]


def claims(token: str) -> dict[str, Any]:
    parts = token.split(".")
    if len(parts) != 3:
        return {}
    try:
        data = base64.urlsafe_b64decode(parts[1] + "=" * (-len(parts[1]) % 4))
        value = json.loads(data)
    except (ValueError, json.JSONDecodeError):
        return {}
    return value if isinstance(value, dict) else {}


# --------------------------------------------------------------------------- HTTP

class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


_OPENER = urllib.request.build_opener(_NoRedirect())
_SAFE_HEADERS = ("content-type", "retry-after", "x-owa-error", "x-owa-returncode", "request-id")


@dataclass
class Resp:
    status: int | None
    payload: Any = None
    raw: bytes = b""
    headers: dict[str, str] = field(default_factory=dict)
    error_code: str | None = None
    elapsed_ms: int = 0
    capped: bool = False

    @property
    def ok(self) -> bool:
        return self.status is not None and 200 <= self.status < 300

    def summary(self) -> dict[str, Any]:
        out: dict[str, Any] = {"http_status": self.status, "elapsed_ms": self.elapsed_ms}
        if self.error_code:
            out["error_code"] = self.error_code
        if self.capped:
            out["response_capped"] = True
        for name in ("x-owa-error", "x-owa-returncode"):
            if name in self.headers:
                out[name] = self.headers[name][:80]
        return out


def _error_code(raw: bytes) -> str | None:
    try:
        data = json.loads(raw.decode("utf-8"))
    except (UnicodeError, json.JSONDecodeError):
        return None
    if not isinstance(data, dict):
        return None
    err = data.get("error")
    if isinstance(err, dict) and isinstance(err.get("code"), str):
        return err["code"][:80]
    if isinstance(err, str):
        codes = data.get("error_codes")
        return f"{err[:60]}{' AADSTS' + str(codes[0]) if isinstance(codes, list) and codes else ''}"
    for key in ("ErrorCode", "ResponseCode"):
        if isinstance(data.get(key), (str, int)):
            return str(data[key])[:80]
    return None


def http(
    method: str,
    url: str,
    *,
    token: str | None = None,
    json_body: Any = None,
    form: dict[str, str] | None = None,
    headers: dict[str, str] | None = None,
    limit: int = 4 * 1024 * 1024,
    parse_json: bool = True,
    retries_429: int = 1,
) -> Resp:
    host = urllib.parse.urlsplit(url).hostname or ""
    if host not in ALLOWED_HOSTS:
        return Resp(None, error_code=f"host_not_allowed")
    hdrs = {"Accept": "application/json", **(headers or {})}
    data = None
    if token:
        hdrs["Authorization"] = f"Bearer {token}"
    if json_body is not None:
        hdrs["Content-Type"] = "application/json; charset=utf-8"
        data = json.dumps(json_body, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
    elif form is not None:
        hdrs["Content-Type"] = "application/x-www-form-urlencoded"
        data = urllib.parse.urlencode(form).encode("utf-8")
    for attempt in range(retries_429 + 1):
        start = time.monotonic()
        req = urllib.request.Request(url, data=data, headers=hdrs, method=method)
        try:
            with _OPENER.open(req, timeout=60) as r:
                raw = r.read(limit + 1)
                resp = Resp(r.status, raw=raw[:limit], capped=len(raw) > limit)
                resp.headers = {k: r.headers[k] for k in _SAFE_HEADERS if r.headers.get(k)}
        except urllib.error.HTTPError as exc:
            raw = exc.read(256 * 1024)
            resp = Resp(exc.code, raw=raw, error_code=_error_code(raw))
            resp.headers = {k: exc.headers[k] for k in _SAFE_HEADERS if exc.headers.get(k)}
        except (OSError, TimeoutError, urllib.error.URLError) as exc:
            resp = Resp(None, error_code=f"network:{type(exc).__name__}")
        resp.elapsed_ms = int((time.monotonic() - start) * 1000)
        if resp.status == 429 and attempt < retries_429 and method == "GET":
            time.sleep(min(int(resp.headers.get("retry-after", "5") or 5), 30))
            continue
        break
    if parse_json and resp.raw and not resp.capped:
        try:
            resp.payload = json.loads(resp.raw.decode("utf-8-sig"))
        except (UnicodeError, json.JSONDecodeError):
            resp.payload = None
    return resp


# --------------------------------------------------------------------------- auth

class SignInRequired(RuntimeError):
    pass


def _load() -> dict[str, Any]:
    try:
        return json.loads(TOKEN_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def _save(store: dict[str, Any]) -> None:
    LOCAL.mkdir(exist_ok=True)
    tmp = TOKEN_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(store, indent=1), encoding="utf-8")
    os.replace(tmp, TOKEN_FILE)


def _guard(client_id: str, scope: str) -> None:
    if (client_id, "*") in DENIED or (client_id, scope) in DENIED:
        raise SystemExit(f"Refusing recorded AADSTS65002 pair: {client_id} / {scope}")


def _scope_string(scope: str) -> str:
    return f"{scope} offline_access openid profile"


def _account(result: dict[str, Any]) -> dict[str, Any]:
    idc = claims(result.get("id_token", "")) if result.get("id_token") else {}
    acc = claims(result.get("access_token", ""))
    return {
        "upn": idc.get("preferred_username") or acc.get("upn") or acc.get("unique_name"),
        "oid": idc.get("oid") or acc.get("oid"),
        "tid": idc.get("tid") or acc.get("tid"),
    }


def _store_result(profile: str, client_id: str, scope: str, result: dict[str, Any]) -> dict[str, Any]:
    store = _load()
    previous = store.get(profile, {})
    entry = {
        "client_id": client_id,
        "scope": scope,
        "access_token": result["access_token"],
        "expires_at": int(time.time()) + int(result.get("expires_in", 3600)),
        "refresh_token": result.get("refresh_token") or previous.get("refresh_token"),
        "account": _account(result) if result.get("id_token") else previous.get("account") or _account(result),
    }
    store[profile] = entry
    _save(store)
    return entry


def device_code_sign_in(profile: str) -> dict[str, Any]:
    client_id, scope = PROFILES[profile]
    _guard(client_id, scope)
    start = http("POST", f"{AUTHORITY}/devicecode", form={"client_id": client_id, "scope": _scope_string(scope)})
    flow = start.payload if isinstance(start.payload, dict) else {}
    if not flow.get("device_code"):
        raise SystemExit(f"Device-code start failed: {start.summary()}")
    print(f"\n>>> {flow.get('message')}\n", flush=True)
    deadline = time.time() + int(flow.get("expires_in", 900))
    interval = int(flow.get("interval", 5))
    while time.time() < deadline:
        time.sleep(interval)
        res = http("POST", f"{AUTHORITY}/token", form={
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": client_id,
            "device_code": flow["device_code"],
        })
        body = res.payload if isinstance(res.payload, dict) else {}
        if body.get("access_token"):
            return _store_result(profile, client_id, scope, body)
        err = body.get("error")
        if err == "authorization_pending":
            continue
        if err == "slow_down":
            interval += 5
            continue
        raise SystemExit(f"Sign-in failed: {err} {body.get('error_codes')}")
    raise SystemExit("Sign-in timed out.")


def _refresh(profile: str, client_id: str, scope: str, refresh_token: str) -> dict[str, Any] | None:
    res = http("POST", f"{AUTHORITY}/token", form={
        "grant_type": "refresh_token",
        "client_id": client_id,
        "refresh_token": refresh_token,
        "scope": _scope_string(scope),
    })
    body = res.payload if isinstance(res.payload, dict) else {}
    if body.get("access_token"):
        return _store_result(profile, client_id, scope, body)
    if body.get("error"):
        print(f"[auth] refresh for '{profile}' failed: {body.get('error')} {body.get('error_codes')}", file=sys.stderr)
    return None


def get_token(profile: str) -> str:
    """Cached access token -> refresh grant -> refresh with a sibling profile's RT (same client)."""
    client_id, scope = PROFILES[profile]
    _guard(client_id, scope)
    store = _load()
    entry = store.get(profile)
    if entry and entry.get("expires_at", 0) > time.time() + 120:
        return _checked(profile, entry)
    candidates = []
    if entry and entry.get("refresh_token"):
        candidates.append(entry["refresh_token"])
    for other, data in store.items():
        if other != profile and data.get("client_id") == client_id and data.get("refresh_token"):
            candidates.append(data["refresh_token"])
    for rt in candidates:
        fresh = _refresh(profile, client_id, scope, rt)
        if fresh:
            return _checked(profile, fresh)
    raise SignInRequired(f"No usable token for '{profile}'. Run: python research/probes/auth.py {profile}")


def _checked(profile: str, entry: dict[str, Any]) -> str:
    upn = (entry.get("account") or {}).get("upn") or ""
    if upn.casefold() != EXPECTED_USER.casefold():
        raise SystemExit(f"Token for '{profile}' does not belong to the expected account; refusing to use it.")
    return entry["access_token"]


def try_scope(client_id: str, scope: str) -> dict[str, Any]:
    """Ask for an extra resource scope with a stored refresh token of the same client. Nothing is stored."""
    if (client_id, scope) in DENIED:
        return {"result": "skipped_recorded_denial"}
    rts = [d["refresh_token"] for d in _load().values() if d.get("client_id") == client_id and d.get("refresh_token")]
    if not rts:
        return {"result": "no_refresh_token_for_client"}
    res = http("POST", f"{AUTHORITY}/token", form={
        "grant_type": "refresh_token", "client_id": client_id,
        "refresh_token": rts[0], "scope": _scope_string(scope)})
    body = res.payload if isinstance(res.payload, dict) else {}
    if body.get("access_token"):
        c = claims(body["access_token"])
        scp = (c.get("scp") or "").split()
        return {"result": "granted", "audience": c.get("aud"), "scope_count": len(scp),
                "mail_scopes": sorted(x for x in scp if x.startswith("Mail."))}
    codes = body.get("error_codes") or []
    return {"result": "denied", "error": body.get("error"), "aadsts": [f"AADSTS{c}" for c in codes]}


def account(profile: str) -> dict[str, Any]:
    return (_load().get(profile) or {}).get("account") or {}


def token_status() -> dict[str, Any]:
    out = {}
    for profile, entry in _load().items():
        c = claims(entry.get("access_token", ""))
        scp = (c.get("scp") or "").split()
        out[profile] = {
            "client_id": entry.get("client_id"),
            "audience": c.get("aud"),
            "scope_count": len(scp),
            "mail_scopes": sorted(s for s in scp if s.startswith("Mail.") or "Search" in s),
            "expires_in_s": int(entry.get("expires_at", 0) - time.time()),
            "refresh_token_present": bool(entry.get("refresh_token")),
            "account_matches_expected": ((entry.get("account") or {}).get("upn") or "").casefold() == EXPECTED_USER.casefold(),
        }
    return out


# --------------------------------------------------------------------------- Graph

def graph_url(path: str, **params: Any) -> str:
    if path.startswith("https://"):
        return path
    query = urllib.parse.urlencode({k: v for k, v in params.items() if v is not None}, quote_via=urllib.parse.quote)
    return f"{GRAPH}{path}" + (f"?{query}" if query else "")


def graph(token: str, path: str, *, method: str = "GET", body: Any = None, prefer: str | None = IMMUTABLE,
          headers: dict[str, str] | None = None, **params: Any) -> Resp:
    hdrs = dict(headers or {})
    if prefer:
        hdrs["Prefer"] = prefer
    return http(method, graph_url(path, **params), token=token, json_body=body, headers=hdrs)


def graph_pages(token: str, path: str, *, max_pages: int = 10, prefer: str | None = IMMUTABLE,
                headers: dict[str, str] | None = None, **params: Any) -> Iterator[Resp]:
    url: str | None = graph_url(path, **params)
    pages = 0
    while url and pages < max_pages:
        r = graph(token, url, prefer=prefer, headers=headers)
        pages += 1
        yield r
        if not r.ok or not isinstance(r.payload, dict):
            return
        nxt = r.payload.get("@odata.nextLink")
        url = nxt if isinstance(nxt, str) and nxt.startswith(GRAPH) else None


def rows(r: Resp) -> list[dict[str, Any]]:
    v = r.payload.get("value") if isinstance(r.payload, dict) else None
    return [x for x in v if isinstance(x, dict)] if isinstance(v, list) else []


WELL_KNOWN = ["inbox", "sentitems", "drafts", "outbox", "deleteditems", "junkemail", "archive",
              "recoverableitemsdeletions", "searchfolders", "conversationhistory", "syncissues"]


def well_known_ids(token: str) -> dict[str, str]:
    """Map folder id -> well-known alias for the aliases that resolve."""
    out = {}
    for alias in WELL_KNOWN:
        r = graph(token, f"/me/mailFolders/{alias}", **{"$select": "id"})
        if r.ok and isinstance(r.payload, dict) and r.payload.get("id"):
            out[r.payload["id"]] = alias
    return out


# --------------------------------------------------------------------------- OWS (Outlook Web JSON RPC)

def _cv() -> str:
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    return "".join(secrets.choice(alphabet) for _ in range(22)) + ".0"


def ows(token: str, action: str, body_type: str, body: dict[str, Any], *,
        version: str = "V2018_01_08", anchor: str = EXPECTED_USER) -> Resp:
    """Bearer-only OWS call using the envelope proven in the 2026-10-01 probes."""
    envelope = {
        "__type": f"{action}JsonRequest:#Exchange",
        "Header": {"__type": "JsonRequestHeaders:#Exchange", "RequestServerVersion": version},
        "Body": {"__type": f"{body_type}:#Exchange", **body},
    }
    serialized = json.dumps(envelope, separators=(",", ":"), ensure_ascii=False)
    encoded = urllib.parse.quote(serialized, safe="-_.!~*'()")
    headers = {
        "Action": action,
        "X-OWA-ActionSource": action,
        "X-OWA-SessionId": _cv(),
        "X-OWA-CorrelationId": str(uuid.uuid4()),
        "MS-CV": _cv(),
        "Prefer": IMMUTABLE,
        "X-AnchorMailbox": f"AAD-SMTP:{anchor}",
    }
    url = f"{OWS_URL}?action={urllib.parse.quote(action)}&app=Mail"
    if len(encoded) <= 2048:
        headers["X-OWA-UrlPostData"] = encoded
        headers["Content-Type"] = "application/json; charset=utf-8"
        return http("POST", url, token=token, headers=headers, retries_429=0)
    return http("POST", url, token=token, json_body=envelope, headers=headers, retries_429=0)


def ows_items(r: Resp) -> list[dict[str, Any]]:
    body = r.payload.get("Body") if isinstance(r.payload, dict) else None
    msgs = body.get("ResponseMessages") if isinstance(body, dict) else None
    items = msgs.get("Items") if isinstance(msgs, dict) else None
    return [x for x in items if isinstance(x, dict)] if isinstance(items, list) else []


def ows_result(r: Resp) -> dict[str, Any]:
    items = ows_items(r)
    first = items[0] if items else {}
    return {**r.summary(), "response_class": first.get("ResponseClass"), "response_code": first.get("ResponseCode")}


def ows_message(subject: str, text: str, disposition: str) -> dict[str, Any]:
    """CreateItem body proven by the 2026-10-01 self-send (disposition SendAndSaveCopy).

    ``SaveOnly`` creates a draft without sending.
    """
    me = {"__type": "EmailAddress:#Exchange", "EmailAddress": EXPECTED_USER, "RoutingType": "SMTP"}
    message = {
        "__type": "Message:#Exchange",
        "Body": {"__type": "BodyContentType:#Exchange", "BodyType": "Text", "Value": text},
        "From": {}, "ToRecipients": [me], "CcRecipients": [], "BccRecipients": [],
        "Subject": subject, "Importance": "Normal",
        "IsDeliveryReceiptRequested": False, "IsReadReceiptRequested": False, "IsSendIndividually": False,
        "MessageDisposition": disposition, "ShouldIgnoreChangeKey": True, "operation": "New",
    }
    return {
        "ClientSupportsIrm": True, "ComposeOperation": "newMail", "MessageDisposition": disposition,
        "Items": [message], "TimeFormat": "", "SendOnNotFoundError": True, "RemoteExecute": True,
        "ShouldSuppressReadReceipt": True, "OutboundCharset": "AutoDetect",
        "ItemShape": {"__type": "ItemResponseShape:#Exchange", "BaseShape": "IdOnly", "ClientSupportsIrm": True,
                      "AdditionalProperties": [{"__type": "PropertyUri:#Exchange", "FieldURI": "ItemLastModifiedTime"}]},
        "ShapeName": "MailCompose",
    }


def need(profile: str) -> str:
    try:
        return get_token(profile)
    except SignInRequired as exc:
        emit({"result": "sign_in_required", "detail": str(exc)})
        raise SystemExit(3)
