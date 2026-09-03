"""LLM client abstraction: turns raw book text into structured scenes, panels, and a character bible."""
from __future__ import annotations

import json
from abc import ABC, abstractmethod
from typing import Any

from book_to_comic.models import Character, Scene, Panel

SEGMENTATION_SYSTEM_PROMPT = """You are adapting a book into a comic script.
Given a section of a book's text, break it into visual scenes and, within each scene, into
individual comic panels (2-6 panels per scene). For each panel, write a concise visual
description suitable for an image generation prompt (setting, action, framing, mood) and,
if applicable, a short line of dialogue or a caption.

Also list every named character who appears in this section along with a canonical physical
description (age, build, hair, clothing, distinguishing features) inferred or stated in the
text. Keep descriptions stable and specific so the same character can be drawn consistently
across many panels.

Respond ONLY with JSON matching this shape:
{
  "characters": [{"name": str, "description": str}],
  "scenes": [
    {
      "summary": str,
      "setting": str,
      "character_names": [str],
      "panels": [
        {"prompt": str, "character_names": [str], "dialogue": str|null, "caption": str|null}
      ]
    }
  ]
}
"""


class LLMClient(ABC):
    @abstractmethod
    def complete_json(self, system: str, user: str) -> dict[str, Any]:
        """Send a prompt and return the parsed JSON response."""


class AnthropicLLMClient(LLMClient):
    def __init__(self, model: str = "claude-sonnet-5", api_key: str | None = None):
        import anthropic

        self._client = anthropic.Anthropic(api_key=api_key)
        self._model = model

    def complete_json(self, system: str, user: str) -> dict[str, Any]:
        response = self._client.messages.create(
            model=self._model,
            max_tokens=4096,
            system=system,
            messages=[{"role": "user", "content": user}],
        )
        text = "".join(block.text for block in response.content if block.type == "text")
        return _extract_json(text)


def _extract_json(text: str) -> dict[str, Any]:
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1:
        raise ValueError(f"No JSON object found in LLM response: {text[:200]!r}")
    return json.loads(text[start : end + 1])


def segment_section(client: LLMClient, section_text: str, scene_index_offset: int) -> tuple[list[Character], list[Scene]]:
    """Ask the LLM to split one chunk of book text into scenes/panels and extract character descriptions."""
    data = client.complete_json(SEGMENTATION_SYSTEM_PROMPT, section_text)

    characters = [Character(name=c["name"], description=c["description"]) for c in data.get("characters", [])]

    scenes: list[Scene] = []
    for i, raw_scene in enumerate(data.get("scenes", [])):
        panels = [
            Panel(
                scene_index=scene_index_offset + i,
                panel_index=j,
                prompt=p["prompt"],
                character_names=p.get("character_names", []),
                dialogue=p.get("dialogue"),
                caption=p.get("caption"),
            )
            for j, p in enumerate(raw_scene.get("panels", []))
        ]
        scenes.append(
            Scene(
                index=scene_index_offset + i,
                summary=raw_scene.get("summary", ""),
                setting=raw_scene.get("setting", ""),
                character_names=raw_scene.get("character_names", []),
                panels=panels,
            )
        )
    return characters, scenes


def merge_character_descriptions(all_characters: list[Character]) -> dict[str, Character]:
    """Collapse repeated character mentions across sections into one canonical entry per name.

    First description seen wins; this keeps the bible stable instead of letting later, possibly
    inconsistent, re-descriptions overwrite the anchor text used for every prompt.
    """
    bible: dict[str, Character] = {}
    for character in all_characters:
        if character.name not in bible:
            bible[character.name] = character
    return bible
