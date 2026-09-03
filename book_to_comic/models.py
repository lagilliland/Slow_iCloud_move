"""Shared data structures for the book-to-comic pipeline."""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class Character:
    name: str
    description: str  # canonical appearance text reused in every prompt ("the character bible")
    reference_image_path: str | None = None  # anchor portrait used for reference-conditioned generation
    seed: int | None = None  # base seed reused across this character's panels, where the backend supports it


@dataclass
class Panel:
    scene_index: int
    panel_index: int
    prompt: str
    character_names: list[str]
    dialogue: str | None = None
    caption: str | None = None
    image_path: str | None = None
    retries: int = 0


@dataclass
class Scene:
    index: int
    summary: str
    setting: str
    character_names: list[str]
    panels: list[Panel] = field(default_factory=list)


@dataclass
class ComicPage:
    index: int
    panel_image_paths: list[str]
    output_path: str | None = None
