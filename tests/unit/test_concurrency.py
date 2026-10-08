from __future__ import annotations

import asyncio

import pytest

from outlook_connector.service.concurrency import gather_cancel_on_error


async def test_cancelling_caller_waits_for_all_read_cleanup() -> None:
    started, cleaned_up = asyncio.Event(), asyncio.Event()

    async def read() -> None:
        started.set()
        try:
            await asyncio.Event().wait()
        finally:
            await asyncio.sleep(0)
            cleaned_up.set()

    caller = asyncio.create_task(gather_cancel_on_error(read()))
    await started.wait()
    caller.cancel()
    with pytest.raises(asyncio.CancelledError):
        await caller
    assert cleaned_up.is_set()
