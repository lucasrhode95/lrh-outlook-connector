"""Ordered body parts extracted from source structure before plain-text rendering."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class TextPart:
    text: str


@dataclass(frozen=True)
class ImagePart:
    content_id: str | None
    alt: str = ""


BodyPart = TextPart | ImagePart
