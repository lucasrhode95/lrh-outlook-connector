"""Composition root: config → tokens → transport → reader → store → services.

Built lazily on first use, so a starting process does no network or disk work (architecture §2).
"""

from __future__ import annotations

from dataclasses import dataclass

import httpx

from outlook_connector.auth.tokens import Account, TokenProvider
from outlook_connector.domain.errors import AuthenticationRequired
from outlook_connector.remote.graph import Graph
from outlook_connector.remote.graph_mail import GraphMailReader
from outlook_connector.remote.transport import Transport
from outlook_connector.service.export.orchestrator import Exports
from outlook_connector.service.files import Files
from outlook_connector.service.mailbox import Mailbox
from outlook_connector.service.threads import Threads
from outlook_connector.store.db import Store, store_path


@dataclass
class Services:
    account: Account
    mailbox: Mailbox
    threads: Threads
    exports: Exports
    files: Files


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
            threads = Threads(mailbox)
            self._services = Services(account, mailbox, threads, Exports(threads), Files(mailbox))
        return self._services

    async def aclose(self) -> None:
        if self._transport is not None:
            await self._transport.aclose()
