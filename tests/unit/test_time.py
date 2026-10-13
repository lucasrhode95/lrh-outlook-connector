from __future__ import annotations

from datetime import UTC, datetime, timedelta, timezone

from outlook_connector.domain.models import ExportRequest
from outlook_connector.domain.time import as_utc, optional_utc


def test_a_time_without_a_zone_is_utc() -> None:
    assert as_utc(datetime(2026, 10, 1, 9)) == datetime(2026, 10, 1, 9, tzinfo=UTC)


def test_an_aware_time_is_converted_to_utc() -> None:
    local = datetime(2026, 10, 1, 9, tzinfo=timezone(timedelta(hours=-3)))
    assert as_utc(local) == datetime(2026, 10, 1, 12, tzinfo=UTC)


def test_optional_utc_keeps_none() -> None:
    assert optional_utc(None) is None


def test_export_request_normalizes_its_window_once() -> None:
    request = ExportRequest(
        since=datetime(2026, 10, 1, 9),
        until=datetime(2026, 10, 2, 9, tzinfo=timezone(timedelta(hours=2))),
    )
    assert request.since == datetime(2026, 10, 1, 9, tzinfo=UTC)
    assert request.until == datetime(2026, 10, 2, 7, tzinfo=UTC)
