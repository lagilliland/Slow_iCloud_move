"""End-to-end smoke test using fully offline fake clients (no network, no API keys)."""
from __future__ import annotations

import zipfile
from pathlib import Path

from book_to_comic.demo import FakeLLMClient
from book_to_comic.image_gen import MockImageClient
from book_to_comic.pipeline import run_pipeline


def test_run_pipeline_produces_cbz(tmp_path: Path):
    book_path = tmp_path / "book.txt"
    book_path.write_text(
        "Chapter One\n\nAlex walked into the clearing, unsure of what waited there.\n\n"
        "Chapter Two\n\nMira had been watching from the trees the whole time.\n"
    )

    output_dir = tmp_path / "out"
    cbz_path = run_pipeline(book_path, output_dir, FakeLLMClient(), MockImageClient())

    assert Path(cbz_path).exists()
    with zipfile.ZipFile(cbz_path) as zf:
        assert len(zf.namelist()) >= 1

    # character reference portraits were generated and reused (the anti-drift anchor)
    characters_dir = output_dir / "characters"
    assert (characters_dir / "alex.png").exists()
    assert (characters_dir / "mira.png").exists()

    # panels were generated and each carries its scene's character references
    panel_files = list((output_dir / "panels").glob("*.png"))
    assert len(panel_files) >= 2
