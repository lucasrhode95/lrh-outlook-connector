"""Graph ↔ OWS ids (research §4.2), the only place that converts them.

A Graph immutable REST id is URL-safe base64; OWS takes the same id in standard base64
(``-``→``/``, ``_``→``+``). Conversation ids map the same way. Verified live 2026-10-02.
"""

from __future__ import annotations

_TO_OWS = str.maketrans("-_", "/+")
_TO_GRAPH = str.maketrans("/+", "-_")


def to_ows(graph_id: str) -> str:
    return graph_id.translate(_TO_OWS)


def to_graph(ows_id: str) -> str:
    return ows_id.translate(_TO_GRAPH)
