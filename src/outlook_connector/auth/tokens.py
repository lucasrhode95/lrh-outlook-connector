"""One centralized MSAL token provider serving every named client profile.

- One token cache file for all profiles (MSAL keys entries by client id), one account.
- Encrypted persistence by default; fail closed. ``unsecure=True`` selects a separate
  plaintext development cache and is never chosen implicitly.
- Silent acquisition only, except ``sign_in``, which the ``auth`` command alone calls.
- A cross-process lock around each acquisition, so concurrent processes do not race on
  refresh-token rotation.
- Every profile must resolve to the same Microsoft account (tenant id + object id).
"""

from __future__ import annotations

import base64
import hashlib
import json
import time
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Any

import msal
from msal_extensions import CrossPlatLock, FilePersistence, PersistedTokenCache, build_encrypted_persistence

from outlook_connector import config
from outlook_connector.domain.errors import (
    AccountMismatch,
    AuthenticationRequired,
    ConfigurationError,
    SecureStorageUnavailable,
    Upstream,
)

# MSAL error codes meaning "the stored credentials are no longer accepted".
_REJECTED = {"invalid_grant", "interaction_required", "login_required", "consent_required", "access_denied"}
_NETWORK_RETRIES = 3
_MEMO_MARGIN_SECONDS = 300  # reuse an in-memory access token until 5 minutes before it expires

AppFactory = Callable[..., Any]


@dataclass(frozen=True)
class Account:
    tenant_id: str
    object_id: str
    username: str | None

    @property
    def fingerprint(self) -> str:
        """Privacy-preserving, stable identity for binding local data to this account."""
        return hashlib.sha256(f"{self.tenant_id}:{self.object_id}".encode()).hexdigest()[:16]

    @classmethod
    def from_msal(cls, account: dict[str, Any]) -> Account:
        object_id, _, tenant_id = str(account.get("home_account_id", "")).partition(".")
        return cls(tenant_id=tenant_id, object_id=object_id, username=account.get("username"))


@dataclass(frozen=True)
class AccessToken:
    profile: str
    value: str = field(repr=False)
    source: str  # "cache" or "identity_provider"
    expires_on: int | None

    def claims(self) -> dict[str, Any]:
        return decode_claims(self.value)


@dataclass(frozen=True)
class ProfileStatus:
    profile: str
    client_id: str
    scopes: tuple[str, ...]
    purpose: str
    signed_in: bool  # a refresh token for this client is cached
    access_token_expires_on: int | None


@dataclass(frozen=True)
class CacheStatus:
    mode: str
    path: Path
    exists: bool
    accounts: tuple[Account, ...]
    profiles: tuple[ProfileStatus, ...]


def decode_claims(token: str) -> dict[str, Any]:
    """Decode a JWT payload without verification (inspection only)."""
    parts = token.split(".")
    if len(parts) != 3:
        return {}
    try:
        payload = json.loads(base64.urlsafe_b64decode(parts[1] + "=" * (-len(parts[1]) % 4)))
    except (ValueError, json.JSONDecodeError):
        return {}
    return payload if isinstance(payload, dict) else {}


class TokenProvider:
    def __init__(
        self,
        *,
        unsecure: bool = False,
        cache_path: Path | None = None,
        profiles: dict[str, config.TokenProfile] | None = None,
        authority: str = config.AUTHORITY,
        app_factory: AppFactory | None = None,
    ) -> None:
        self.unsecure = unsecure
        self.cache_path = cache_path or config.token_cache_path(unsecure=unsecure)
        self.profiles = profiles or config.PROFILES
        self._authority = authority
        self._app_factory = app_factory or msal.PublicClientApplication
        # Per process: building an MSAL app costs a network round trip, so build each once, and keep
        # the current access token in memory instead of re-reading the locked cache on every request.
        self._apps: dict[str, Any] = {}
        self._token_cache: PersistedTokenCache | None = None
        self._memo: dict[str, AccessToken] = {}
        for profile in self.profiles.values():
            denied = [s for s in profile.scopes if (profile.client_id, s) in config.DENIED_PAIRS]
            if denied:
                raise ConfigurationError(
                    f"Profile '{profile.name}' requests {denied}, which Microsoft denied for this client "
                    "(AADSTS65002). See docs/outlook-api-research.md §2."
                )

    # ------------------------------------------------------------------ public API

    @property
    def mode(self) -> str:
        return "plaintext-development" if self.unsecure else "encrypted"

    def sign_in_command(self, profile: str) -> str:
        return f"{config.CLI_NAME} auth {profile}" + (" --unsecure" if self.unsecure else "")

    def get_token(
        self, profile: str, *, force_refresh: bool = False, claims_challenge: str | None = None
    ) -> AccessToken:
        """Silent acquisition (memory, then cache, refreshing if needed). Never prompts.

        ``force_refresh`` (after the service rejected a token with 401) skips the in-memory and cached
        access tokens; ``claims_challenge`` passes the service's continuous-access-evaluation challenge.
        """
        spec = self._profile(profile)
        memo = self._memo.get(profile)
        reuse = not force_refresh and claims_challenge is None
        if reuse and memo and memo.expires_on and memo.expires_on - time.time() > _MEMO_MARGIN_SECONDS:
            return replace(memo, source="memory")
        self._memo.pop(profile, None)
        with self._lock():
            app = self._retry(lambda: self._app(spec))
            account = self._bound_account(app.get_accounts())
            if account is None:
                raise AuthenticationRequired(
                    f"No Microsoft sign-in found for '{profile}'.", command=self.sign_in_command(profile)
                )
            options: dict[str, Any] = {}
            if force_refresh:
                options["force_refresh"] = True
            if claims_challenge:
                options["claims_challenge"] = claims_challenge
            result = self._retry(
                lambda: app.acquire_token_silent_with_error(list(spec.scopes), account=account, **options)
            )
        token = self._token_from_result(profile, result)
        self._memo[profile] = token
        return token

    def sign_in(self, profile: str, *, show: Callable[[str], None] = print) -> AccessToken:
        """Device-code sign-in for one profile. Only the `auth` command calls this."""
        spec = self._profile(profile)
        app = self._retry(lambda: self._app(spec))
        existing = self._bound_account(app.get_accounts())
        flow = self._retry(lambda: app.initiate_device_flow(scopes=list(spec.scopes)))
        if not isinstance(flow, dict) or "user_code" not in flow:
            raise Upstream(f"Microsoft did not start device-code sign-in: {_error_text(flow)}")
        show(str(flow.get("message", "")))
        result = app.acquire_token_by_device_flow(flow)  # polls until done or expired
        if isinstance(result, dict) and result.get("error") in ("authorization_pending", "expired_token"):
            raise AuthenticationRequired(
                "The sign-in code expired before it was used.", command=self.sign_in_command(profile)
            )
        token = self._token_from_result(profile, result)
        signed_in = _account_from_result(result)
        if (
            existing is not None
            and signed_in is not None
            and (
                (signed_in.tenant_id, signed_in.object_id)
                != (Account.from_msal(existing).tenant_id, Account.from_msal(existing).object_id)
            )
        ):
            for cached in app.get_accounts():
                if Account.from_msal(cached).object_id == signed_in.object_id:
                    app.remove_account(cached)
            raise AccountMismatch(
                f"You signed in as a different Microsoft account than the one already bound "
                f"({Account.from_msal(existing).username}). Run `{config.CLI_NAME} auth --sign-out` first "
                "to switch accounts."
            )
        self._memo[profile] = token
        return token

    def sign_out(self) -> bool:
        """Delete the selected token cache. Returns whether a cache existed."""
        existed = self.cache_path.exists()
        self._apps.clear()
        self._memo.clear()
        self._token_cache = None
        with self._lock():
            for path in (self.cache_path, Path(f"{self.cache_path}.lockfile")):
                path.unlink(missing_ok=True)
        return existed

    def status(self) -> CacheStatus:
        """Offline view of the cache: no network calls, no token values."""
        exists = self.cache_path.exists()
        cache = self._cache() if exists else msal.TokenCache()
        accounts = tuple(Account.from_msal(a) for a in cache.search(msal.TokenCache.CredentialType.ACCOUNT))
        now = time.time()
        rows = []
        for spec in self.profiles.values():
            refresh = list(
                cache.search(
                    msal.TokenCache.CredentialType.REFRESH_TOKEN, query={"client_id": spec.client_id}
                )
            )
            expiries = [
                int(at.get("expires_on", 0))
                for at in cache.search(
                    msal.TokenCache.CredentialType.ACCESS_TOKEN, query={"client_id": spec.client_id}
                )
                if int(at.get("expires_on", 0)) > now
            ]
            rows.append(
                ProfileStatus(
                    spec.name,
                    spec.client_id,
                    spec.scopes,
                    spec.purpose,
                    signed_in=bool(refresh),
                    access_token_expires_on=max(expiries) if expiries else None,
                )
            )
        return CacheStatus(self.mode, self.cache_path, exists, accounts, tuple(rows))

    # ------------------------------------------------------------------ internals

    def _profile(self, name: str) -> config.TokenProfile:
        try:
            return self.profiles[name]
        except KeyError:
            raise ConfigurationError(
                f"Unknown token profile '{name}'. Known: {sorted(self.profiles)}"
            ) from None

    def _app(self, spec: config.TokenProfile) -> Any:
        if spec.name not in self._apps:
            if self._token_cache is None:
                self._token_cache = self._cache()  # re-reads the file when other processes change it
            self._apps[spec.name] = self._app_factory(
                spec.client_id, authority=self._authority, token_cache=self._token_cache
            )
        return self._apps[spec.name]

    def _cache(self) -> PersistedTokenCache:
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        if self.unsecure:
            return PersistedTokenCache(FilePersistence(str(self.cache_path)))
        try:
            persistence = build_encrypted_persistence(str(self.cache_path))
        except Exception as exc:  # platform keyring/DPAPI unavailable
            raise SecureStorageUnavailable(
                f"Encrypted token storage is unavailable ({type(exc).__name__}). "
                "No plaintext cache will be used. For development only, add --unsecure."
            ) from None
        return PersistedTokenCache(persistence)

    @contextmanager
    def _lock(self) -> Iterator[None]:
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        with CrossPlatLock(f"{self.cache_path}.acquire.lock"):
            yield

    def _bound_account(self, accounts: list[dict[str, Any]]) -> dict[str, Any] | None:
        identities = {(Account.from_msal(a).tenant_id, Account.from_msal(a).object_id) for a in accounts}
        if len(identities) > 1:
            raise AccountMismatch(
                "The token cache holds more than one Microsoft account. "
                f"Run `{config.CLI_NAME} auth --sign-out` and sign in again with one account."
            )
        return accounts[0] if accounts else None

    def _token_from_result(self, profile: str, result: Any) -> AccessToken:
        if isinstance(result, dict) and result.get("access_token"):
            expires_in = result.get("expires_in")
            return AccessToken(
                profile=profile,
                value=result["access_token"],
                source=str(result.get("token_source", "cache")),
                expires_on=int(time.time() + int(expires_in)) if expires_in else None,
            )
        error = str(result.get("error", "")) if isinstance(result, dict) else ""
        if not result:
            raise AuthenticationRequired(
                f"No sign-in yet for '{profile}'.", command=self.sign_in_command(profile)
            )
        if error in _REJECTED:
            raise AuthenticationRequired(
                f"Microsoft no longer accepts the stored sign-in for '{profile}' ({_error_text(result)}).",
                command=self.sign_in_command(profile),
            )
        raise Upstream(f"Token request for '{profile}' failed: {_error_text(result)}")

    @staticmethod
    def _retry(operation: Callable[[], Any]) -> Any:
        """Retry transient network failures. Token requests are idempotent."""
        import requests  # MSAL's HTTP stack

        for attempt in range(_NETWORK_RETRIES):
            try:
                return operation()
            except (requests.ConnectionError, requests.Timeout) as exc:
                if attempt == _NETWORK_RETRIES - 1:
                    raise Upstream(f"Could not reach Microsoft sign-in ({type(exc).__name__}).") from None
                time.sleep(2**attempt)
        raise AssertionError("unreachable")


def _account_from_result(result: dict[str, Any]) -> Account | None:
    claims = result.get("id_token_claims") or {}
    if not claims.get("oid") or not claims.get("tid"):
        return None
    return Account(
        tenant_id=claims["tid"], object_id=claims["oid"], username=claims.get("preferred_username")
    )


def _error_text(result: Any) -> str:
    if not isinstance(result, dict):
        return "no response"
    codes = result.get("error_codes") or []
    suffix = f" AADSTS{codes[0]}" if codes else ""
    return f"{result.get('error', 'unknown_error')}{suffix}"
