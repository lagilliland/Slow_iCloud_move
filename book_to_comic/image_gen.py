"""Image generation client abstraction.

The anti-drift strategy lives mostly here: every panel that includes a character is generated
with that character's reference portrait passed in as an image input (reference/image-to-image
conditioning), not just a text description. Text descriptions alone are not enough to keep a
face, hairstyle, or outfit consistent across dozens of independently-sampled images.
"""
from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path

STYLE_SUFFIX = (
    ", consistent comic book art style, clean bold inked linework, flat cel shading, "
    "muted limited color palette, dynamic panel composition"
)


class ImageClient(ABC):
    @abstractmethod
    def generate(self, prompt: str, out_path: str | Path, seed: int | None = None) -> str:
        """Generate a standalone image (used for character reference portraits) and return its path."""

    @abstractmethod
    def generate_with_references(
        self,
        prompt: str,
        reference_image_paths: list[str],
        out_path: str | Path,
        seed: int | None = None,
    ) -> str:
        """Generate an image conditioned on one or more reference images (character-consistent panel)."""


class OpenAIImageClient(ImageClient):
    """Uses OpenAI's image API. Reference conditioning is done via images.edit with the
    reference portrait(s) as input images, which keeps the depicted character close to the
    anchor image instead of re-imagining it from text alone each time."""

    def __init__(self, model: str = "gpt-image-1", api_key: str | None = None):
        from openai import OpenAI

        self._client = OpenAI(api_key=api_key)
        self._model = model

    def generate(self, prompt: str, out_path: str | Path, seed: int | None = None) -> str:
        result = self._client.images.generate(model=self._model, prompt=prompt + STYLE_SUFFIX, size="1024x1024")
        return _save_b64(result.data[0].b64_json, out_path)

    def generate_with_references(
        self,
        prompt: str,
        reference_image_paths: list[str],
        out_path: str | Path,
        seed: int | None = None,
    ) -> str:
        files = [open(p, "rb") for p in reference_image_paths]
        try:
            result = self._client.images.edit(
                model=self._model,
                image=files,
                prompt=prompt + STYLE_SUFFIX,
                size="1024x1024",
            )
        finally:
            for f in files:
                f.close()
        return _save_b64(result.data[0].b64_json, out_path)


class MockImageClient(ImageClient):
    """Offline stand-in used for local development and tests: draws a placeholder image with
    the prompt text and a visible marker per referenced character, so the pipeline can be
    exercised end to end without network access or API keys."""

    def __init__(self, size: tuple[int, int] = (768, 768)):
        self._size = size

    def generate(self, prompt: str, out_path: str | Path, seed: int | None = None) -> str:
        return self._render(prompt, [], out_path, seed)

    def generate_with_references(
        self,
        prompt: str,
        reference_image_paths: list[str],
        out_path: str | Path,
        seed: int | None = None,
    ) -> str:
        return self._render(prompt, reference_image_paths, out_path, seed)

    def _render(self, prompt: str, reference_image_paths: list[str], out_path: str | Path, seed: int | None) -> str:
        import hashlib
        import textwrap

        from PIL import Image, ImageDraw

        seed_val = seed if seed is not None else int(hashlib.sha256(prompt.encode()).hexdigest(), 16) % (2**32)
        color = (
            50 + (seed_val * 37) % 180,
            50 + (seed_val * 59) % 180,
            50 + (seed_val * 83) % 180,
        )
        img = Image.new("RGB", self._size, color=color)
        draw = ImageDraw.Draw(img)
        wrapped = textwrap.fill(prompt, width=40)
        draw.text((20, 20), wrapped, fill="white")
        if reference_image_paths:
            draw.text((20, self._size[1] - 30), f"refs: {len(reference_image_paths)}", fill="yellow")
        out_path = Path(out_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        img.save(out_path)
        return str(out_path)


def _save_b64(b64_data: str, out_path: str | Path) -> str:
    import base64

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(base64.b64decode(b64_data))
    return str(out_path)
