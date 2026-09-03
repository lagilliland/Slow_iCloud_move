"""Automated QA pass that catches drift the reference-conditioning step misses.

Reference-image conditioning (see image_gen.py) is the primary defense against drift. This
module is the safety net: after a panel is generated, compare it against the character's
reference portrait with a similarity scorer, and regenerate (with a fresh seed) if it falls
below a threshold, up to a retry cap.

The default scorer uses perceptual hashing (imagehash), which is cheap and dependency-light
and catches gross drift (wrong outfit colors, wrong setting bleeding into the character crop,
a completely different pose/character rendered). It is intentionally not a face-recognition
model: swap in `ClipEmbeddingSimilarityScorer` (or any `SimilarityScorer`) for tighter,
identity-level matching when running with a real image backend and GPU available.
"""
from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path


class SimilarityScorer(ABC):
    @abstractmethod
    def similarity(self, image_path_a: str, image_path_b: str) -> float:
        """Return a similarity score in [0, 1], higher meaning more alike."""


class PerceptualHashSimilarityScorer(SimilarityScorer):
    def similarity(self, image_path_a: str, image_path_b: str) -> float:
        import imagehash
        from PIL import Image

        hash_a = imagehash.phash(Image.open(image_path_a))
        hash_b = imagehash.phash(Image.open(image_path_b))
        max_bits = len(hash_a.hash) ** 2
        distance = hash_a - hash_b
        return 1.0 - (distance / max_bits)


class ClipEmbeddingSimilarityScorer(SimilarityScorer):
    """Higher-fidelity alternative: cosine similarity between CLIP image embeddings.
    Requires `open-clip-torch`; not installed by default because it pulls in torch."""

    def __init__(self, model_name: str = "ViT-B-32", pretrained: str = "openai"):
        import open_clip
        import torch

        self._torch = torch
        self._model, _, self._preprocess = open_clip.create_model_and_transforms(model_name, pretrained=pretrained)
        self._model.eval()

    def _embed(self, image_path: str):
        from PIL import Image

        image = self._preprocess(Image.open(image_path).convert("RGB")).unsqueeze(0)
        with self._torch.no_grad():
            features = self._model.encode_image(image)
        return features / features.norm(dim=-1, keepdim=True)

    def similarity(self, image_path_a: str, image_path_b: str) -> float:
        a, b = self._embed(image_path_a), self._embed(image_path_b)
        return float((a @ b.T).item())


def ensure_consistent_panel(
    image_client,
    scorer: SimilarityScorer,
    prompt: str,
    reference_image_paths: list[str],
    out_path: str | Path,
    seed: int | None,
    threshold: float = 0.55,
    max_retries: int = 2,
) -> tuple[str, int]:
    """Generate a panel, retrying with a bumped seed if it drifts too far from the reference(s).

    Returns (final_image_path, retries_used). If no reference images are supplied (a panel with
    no recurring characters), the similarity check is skipped and the first render is kept.
    """
    base_seed = seed or 0
    path = image_client.generate_with_references(prompt, reference_image_paths, out_path, seed=base_seed)

    if not reference_image_paths:
        return path, 0

    for retry in range(max_retries):
        best_similarity = max(scorer.similarity(path, ref) for ref in reference_image_paths)
        if best_similarity >= threshold:
            return path, retry
        path = image_client.generate_with_references(
            prompt, reference_image_paths, out_path, seed=base_seed + retry + 1
        )

    return path, max_retries
