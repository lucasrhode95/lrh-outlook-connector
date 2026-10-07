"""Composition root: config → tokens → transport → reader / writer → store → services.

Built lazily on first use, so a starting process does no network or disk work (architecture §2).
"""

from __future__ import annotations

from dataclasses import dataclass

import httpx

from outlook_connector.auth.tokens import Account, TokenProvider
from outlook_connector.domain.errors import AuthenticationRequired
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
from outlook_connector.service.writes import Writes
from outlook_connector.store.db import Store, store_path


@dataclass
class Services:
    account: Account
    mailbox: Mailbox
    conversations: Conversations
    exports: Exports
    files: Files
    writes: Writes  # uses the write sign-in, only when a write is made
    mutations: Mutations  # likewise
    rules: Rules  # OWS reads/writes use the bound write sign-in


class AppContext:
    def __init__(
        self,
        *,
        unsecure: bool = False,
        tokens: TokenProvider | None = None,
        http_client: httpx.AsyncClient | None = None,
    ) -> None:
        self.tokens = tokens or TokenProvider(unsecure=unsecure)
        self._http_client = http_client
        self._transport: Transport | None = None
        self._services: Services | None = None

    async def services(self) -> Services:
        if self._services is None:
            claims = self.tokens.get_token("read").claims()
            if not claims.get("tid") or not claims.get("oid"):
                raise AuthenticationRequired(
                    "The read token carries no account identity.", command=self.tokens.sign_in_command("read")
                )
            account = Account(tenant_id=claims["tid"], object_id=claims["oid"], username=claims.get("upn"))
            self._transport = Transport(self.tokens, client=self._http_client)
            reader = GraphMailReader(Graph(self._transport))
            store = Store(store_path(account.fingerprint), account.fingerprint)
            mailbox = Mailbox(reader, store)
            conversations = Conversations(mailbox)
            ows = Ows(self._transport, self.tokens)
            writer = OwsMailWriter(ows)
            writes = Writes(mailbox, writer, account)
            mutations = Mutations(mailbox, writer, writes.check_account)
            self._services = Services(
                account,
                mailbox,
                conversations,
                Exports(conversations),
                Files(mailbox),
                writes,
                mutations,
                Rules(mailbox, OwsRules(ows), account, writes.check_account),
            )
        return self._services

    async def aclose(self) -> None:
        if self._transport is not None:
            await self._transport.aclose()
