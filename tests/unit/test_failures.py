from __future__ import annotations

import pytest

from outlook_connector.domain.errors import Failure, Upstream
from outlook_connector.domain.models import ExportError
from outlook_connector.service.failures import describe, error_from, error_summary, export_error


@pytest.mark.parametrize(
    ("status", "likely", "retry", "fix"),
    [
        (429, "Microsoft throttled the mailbox", True, "export it again in a few minutes"),
        (503, "Microsoft throttled the mailbox", True, "export it again in a few minutes"),
        (500, "Microsoft service or network problem", True, "retry later"),
        (None, "Microsoft service or network problem", True, "retry later"),
        (403, "access denied for this item", False, "retrying will not help"),
        (
            404,
            "not found: may have been deleted, moved out of reach, or the id may be wrong",
            False,
            "refresh and select it again",
        ),
        (400, "unexpected error", False, "report it with the request id"),
    ],
)
def test_one_classification_for_every_failure(status: int | None, likely: str, retry: bool, fix: str) -> None:
    error = export_error("fetching message bodies", Failure(status=status, code="X", request_id="r-1"))
    assert error.likely_cause.startswith(likely) and error.retry is retry and error.fix == fix
    assert error.status == status and error.request_id == "r-1"


def test_no_response_and_errors_without_an_answer() -> None:
    assert describe(
        export_error("listing attachments", Failure(status=None, message="Could not reach x."))
    ) == ("Could not reach x.")
    plain = error_from("downloading an attachment", ValueError("disk full"))
    assert plain.likely_cause == "unexpected error" and plain.message == "disk full" and plain.status is None
    answered = Upstream("flattened text")
    answered.failure = Failure(status=502, code="BadGateway", request_id="r-2")
    assert (
        describe(error_from("downloading an attachment", answered)) == "HTTP 502 BadGateway, request-id r-2"
    )


def test_fetched_and_export_failures_share_one_detail_format() -> None:
    failure = Failure(status=502, code="BadGateway", message="Temporary failure", request_id="r-2")
    exported = ExportError(
        step="fetching message bodies",
        status=502,
        code="BadGateway",
        message="Temporary failure",
        request_id="r-2",
        likely_cause="Microsoft service or network problem",
        retry=True,
        fix="retry later",
    )
    assert describe(failure) == describe(exported)


def test_summary_names_every_kind_of_gap_with_its_causes() -> None:
    throttled = export_error("fetching message bodies", Failure(status=429))
    denied = export_error("fetching message bodies", Failure(status=403))
    broken = export_error("downloading an attachment", Failure(status=500))
    listing = export_error("listing attachments", Failure(status=None))
    assert error_summary([throttled, throttled, denied], [broken], [listing]) == (
        "Export errors: 3 message bodies (2 throttled, 1 access denied), 1 attachment "
        "(1 service or network problem) and the attachments of 1 message (1 service or network problem) "
        "could not be exported; they are marked [EXPORT ERROR] below."
    )
    assert error_summary([], [], []) is None
