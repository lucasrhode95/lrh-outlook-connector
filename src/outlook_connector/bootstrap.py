"""Composition root: config → tokens → transport → reader / writer → store → services.

Built lazily on first use, so a starting process does no network or disk work (architecture §2).
"""

from __future__ import annotations

import asyncio
from dataclasses import dataclass

import httpx
from m365_auth import Account, AccountBoundTokenSource, AuthenticationRequired, TokenProvider

from outlook_connector import config
from outlook_connector.remote.cloud_settings import CloudSettings
from outlook_connector.remote.graph import Graph
from outlook_connector.remote.graph_mail import GraphMailReader
from outlook_connector.remote.ows import Ows
from outlook_connector.remote.ows_mail import OwsMailWriter
from outlook_connector.remote.ows_rules import OwsRules
from outlook_connector.remote.transport import Transport
from outlook_connector.service.conversations import Conversations
from outlook_connector.service.export.orchestrator import Exports
from outlook_connector.service.files import Files
from outlook_connector.service.mailbox import Mailbox
from outlook_connector.service.mutations import Mutations
from outlook_connector.service.rules import Rules
from outlook_connector.service.signatures import Signatures
from outlook_connector.service.writes import Writes
from outlook_connector.store.db import Store, store_path


@dataclass
class Services:
    account: Account
    mailbox: Mailbox
    conversations: Conversations
    exports: Exports
    files: Files
    writes: Writes  # uses the Outlook sign-in, only when a write is made
    mutations: Mutations  # likewise
    rules: Rules  # OWS reads/writes use the bound Outlook sign-in
    signatures: Signatures  # native Cloud Settings use the bound Outlook sign-in


class AppContext:
    def __init__(
        self,
        *,
        unsecure: bool = False,
        tokens: TokenProvider | None = None,
        http_client: httpx.AsyncClient | None = None,
    ) -> None:
        self.tokens = tokens or TokenProvider(unsecure=unsecure, login_command=f"{config.CLI_NAME} auth")
        self._http_client = http_client
        self._transport: Transport | None = None
        self._services: Services | None = None
        self._lock = asyncio.Lock()

    async def services(self) -> Services:
        account = self._current_account()
        if self._services is not None and self._services.account.fingerprint == account.fingerprint:
            return self._services
        async with self._lock:
            account = self._current_account()
            if self._services is None or self._services.account.fingerprint != account.fingerprint:
                self._services = self._build(account)
        return self._services

    def _current_account(self) -> Account:
        claims = self.tokens.get_token("outlook.graph").claims()
        if not claims.get("tid") or not claims.get("oid"):
            raise AuthenticationRequired(
                "The Outlook Graph token carries no account identity.",
                command=self.tokens.sign_in_command("outlook.graph"),
            )
        return Account(tenant_id=str(claims["tid"]), object_id=str(claims["oid"]), username=claims.get("upn"))

    def _build(self, account: Account) -> Services:
        account_tokens = AccountBoundTokenSource(self.tokens, account)
        if self._transport is None:
            self._http_client = self._http_client or httpx.AsyncClient(
                timeout=httpx.Timeout(60.0, connect=15.0), follow_redirects=False
            )
        # Each service set keeps its own immutable account binding. The HTTP connection pool is
        # shared across account switches, but an old service can never inherit the new token source.
        self._transport = Transport(account_tokens, client=self._http_client)
        reader = GraphMailReader(Graph(self._transport))
        store = Store(store_path(account.fingerprint), account.fingerprint)
        mailbox = Mailbox(reader, store)
        conversations = Conversations(mailbox)
        ows = Ows(self._transport, account_tokens)
        writer = OwsMailWriter(ows)
        signatures = Signatures(CloudSettings(self._transport, account_tokens), account)
        writes = Writes(mailbox, writer, account, signatures)
        mutations = Mutations(mailbox, writer, writes.check_account)
        return Services(
            account,
            mailbox,
            conversations,
            Exports(conversations),
            Files(mailbox),
            writes,
            mutations,
            Rules(mailbox, OwsRules(ows), account, writes.check_account),
            signatures,
        )

    async def aclose(self) -> None:
        if self._transport is not None:
            await self._transport.aclose()
