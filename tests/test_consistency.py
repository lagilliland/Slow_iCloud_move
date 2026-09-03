"""Unit test for the drift-retry logic in consistency.ensure_consistent_panel."""
from __future__ import annotations

from pathlib import Path

from book_to_comic.consistency import SimilarityScorer, ensure_consistent_panel
from book_to_comic.image_gen import ImageClient


class _CountingImageClient(ImageClient):
    """Returns a fixed path each call; used to verify retry count/seed behavior in isolation."""

    def __init__(self):
        self.calls: list[int | None] = []

    def generate(self, prompt, out_path, seed=None):
        self.calls.append(seed)
        Path(out_path).write_bytes(b"stub")
        return str(out_path)

    def generate_with_references(self, prompt, reference_image_paths, out_path, seed=None):
        return self.generate(prompt, out_path, seed)


class _ScriptedScorer(SimilarityScorer):
    """Returns a pre-scripted sequence of similarity scores, one per call."""

    def __init__(self, scores: list[float]):
        self._scores = iter(scores)

    def similarity(self, image_path_a: str, image_path_b: str) -> float:
        return next(self._scores)


def test_accepts_first_render_when_similar_enough(tmp_path):
    client = _CountingImageClient()
    scorer = _ScriptedScorer([0.9])

    path, retries = ensure_consistent_panel(
        client, scorer, "prompt", ["ref.png"], tmp_path / "panel.png", seed=5, threshold=0.55
    )

    assert retries == 0
    assert client.calls == [5]


def test_retries_with_bumped_seed_when_drifted(tmp_path):
    client = _CountingImageClient()
    scorer = _ScriptedScorer([0.2, 0.3, 0.9])

    path, retries = ensure_consistent_panel(
        client, scorer, "prompt", ["ref.png"], tmp_path / "panel.png", seed=5, threshold=0.55, max_retries=3
    )

    assert retries == 2
    assert client.calls == [5, 6, 7]


def test_skips_similarity_check_when_no_references(tmp_path):
    client = _CountingImageClient()
    scorer = _ScriptedScorer([])  # would raise StopIteration if called

    path, retries = ensure_consistent_panel(client, scorer, "prompt", [], tmp_path / "panel.png", seed=5)

    assert retries == 0
    assert client.calls == [5]
