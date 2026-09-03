"""Offline fake LLM client: deterministic scene/panel/character JSON without any network calls.

Used by --mock CLI runs and by the test suite to exercise the full pipeline (ingest -> segment
-> character bible -> panel generation -> layout -> export) without API keys.
"""
from __future__ import annotations

from typing import Any

from book_to_comic.llm import LLMClient


class FakeLLMClient(LLMClient):
    def complete_json(self, system: str, user: str) -> dict[str, Any]:
        snippet = user.strip().splitlines()[0][:60] if user.strip() else "the story"
        return {
            "characters": [
                {"name": "Alex", "description": "young adult, short dark hair, green jacket, determined expression"},
                {"name": "Mira", "description": "older woman, silver braid, long blue coat, calm demeanor"},
            ],
            "scenes": [
                {
                    "summary": f"Opening beat: {snippet}",
                    "setting": "a quiet forest clearing at dawn",
                    "character_names": ["Alex", "Mira"],
                    "panels": [
                        {
                            "prompt": "Alex stands at the edge of the clearing, looking uncertain",
                            "character_names": ["Alex"],
                            "dialogue": "Is this really the place?",
                            "caption": None,
                        },
                        {
                            "prompt": "Mira steps out from behind a tree, hand raised in greeting",
                            "character_names": ["Mira"],
                            "dialogue": "You made it further than I expected.",
                            "caption": None,
                        },
                    ],
                }
            ],
        }
