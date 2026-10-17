from __future__ import annotations

from outlook_connector.remote.transport import request_id, service_error


def test_a_nested_error_body_gives_its_code_and_message() -> None:
    body = {"error": {"code": "ErrorItemNotFound", "message": "The item was not found."}}
    assert service_error(body) == ("ErrorItemNotFound", "The item was not found.")


def test_a_flat_error_body_is_read_too() -> None:
    assert service_error({"errorCode": "Throttled", "message": "Slow down."}) == ("Throttled", "Slow down.")


def test_a_body_that_is_not_an_object_gives_nothing() -> None:
    assert service_error(["not", "an", "object"]) == (None, None)


def test_the_request_id_falls_back_to_the_ms_cv_header() -> None:
    assert request_id({"ms-cv": "abc.1"}) == "abc.1"
