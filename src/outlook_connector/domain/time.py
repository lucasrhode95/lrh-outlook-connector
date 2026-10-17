"""The time zone rule for every entry point: a time without a zone is UTC; an aware time becomes UTC.

Entry points that accept times call these once, so no two layers normalize the same value.
"""

from __future__ import annotations

from datetime import UTC, datetime


def as_utc(value: datetime) -> datetime:
    return value.replace(tzinfo=UTC) if value.tzinfo is None else value.astimezone(UTC)


def optional_utc(value: datetime | None) -> datetime | None:
    return None if value is None else as_utc(value)
