"""Concurrent reads that finish or are cancelled before their caller returns."""

from __future__ import annotations

import asyncio
from collections.abc import Awaitable


async def gather_cancel_on_error[T](*awaitables: Awaitable[T]) -> list[T]:
    """Gather results in input order; on failure or cancellation, cancel and await every sibling.

    Callers and the transport bound request concurrency. Preserve the original exception so the
    surfaces can keep mapping ConnectorError to their existing responses without ExceptionGroup.
    """
    tasks = [asyncio.ensure_future(awaitable) for awaitable in awaitables]
    try:
        return await asyncio.gather(*tasks)
    except BaseException:
        for task in tasks:
            task.cancel()
        await asyncio.gather(*tasks, return_exceptions=True)
        raise
