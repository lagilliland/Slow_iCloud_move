"""End-to-end orchestration: book file -> character bible -> panels -> laid-out comic pages."""
from __future__ import annotations

import zipfile
from pathlib import Path

from book_to_comic import ingest, layout
from book_to_comic.consistency import PerceptualHashSimilarityScorer, SimilarityScorer, ensure_consistent_panel
from book_to_comic.image_gen import ImageClient
from book_to_comic.llm import LLMClient, merge_character_descriptions, segment_section
from book_to_comic.models import Character, Scene


def build_character_bible(
    llm_client: LLMClient,
    image_client: ImageClient,
    sections: list[str],
    output_dir: Path,
) -> tuple[dict[str, Character], list[Scene]]:
    """Segment every section, merge character mentions into one canonical entry each, then
    render one reference portrait per character. That portrait is the anchor every later panel
    of that character is conditioned on."""
    all_characters: list[Character] = []
    all_scenes: list[Scene] = []
    for section in sections:
        characters, scenes = segment_section(llm_client, section, scene_index_offset=len(all_scenes))
        all_characters.extend(characters)
        all_scenes.extend(scenes)

    bible = merge_character_descriptions(all_characters)

    portraits_dir = output_dir / "characters"
    for i, (name, character) in enumerate(bible.items()):
        portrait_path = portraits_dir / f"{_slugify(name)}.png"
        character.seed = i * 1000  # stable per-character base seed, reused across all of their panels
        character.reference_image_path = image_client.generate(
            f"character reference portrait sheet, front-facing, neutral pose, {character.description}",
            portrait_path,
            seed=character.seed,
        )

    return bible, all_scenes


def generate_panels(
    image_client: ImageClient,
    bible: dict[str, Character],
    scenes: list[Scene],
    output_dir: Path,
    scorer: SimilarityScorer | None = None,
) -> None:
    scorer = scorer or PerceptualHashSimilarityScorer()
    panels_dir = output_dir / "panels"

    for scene in scenes:
        for panel in scene.panels:
            references = [
                bible[name].reference_image_path
                for name in panel.character_names
                if name in bible and bible[name].reference_image_path
            ]
            seed = bible[panel.character_names[0]].seed if panel.character_names and panel.character_names[0] in bible else None
            full_prompt = f"{scene.setting}. {panel.prompt}"
            out_path = panels_dir / f"scene{scene.index:03d}_panel{panel.panel_index:02d}.png"

            path, retries = ensure_consistent_panel(
                image_client, scorer, full_prompt, references, out_path, seed=seed
            )
            panel.image_path = path
            panel.retries = retries


def lay_out_pages(scenes: list[Scene], output_dir: Path, panels_per_page: int = 4) -> list[str]:
    all_panels = [panel for scene in scenes for panel in scene.panels]
    pages_dir = output_dir / "pages"
    page_paths = []
    for i, page_panels in enumerate(layout.paginate(all_panels, panels_per_page)):
        page = layout.render_page(page_panels, i, pages_dir)
        page_paths.append(page.output_path)
    return page_paths


def export_cbz(page_paths: list[str], output_path: str | Path) -> str:
    output_path = Path(output_path)
    with zipfile.ZipFile(output_path, "w") as zf:
        for path in page_paths:
            zf.write(path, arcname=Path(path).name)
    return str(output_path)


def run_pipeline(
    book_path: str | Path,
    output_dir: str | Path,
    llm_client: LLMClient,
    image_client: ImageClient,
    scorer: SimilarityScorer | None = None,
    panels_per_page: int = 4,
) -> str:
    output_dir = Path(output_dir)
    text = ingest.load_text(book_path)
    sections = ingest.chunk_into_sections(text)

    bible, scenes = build_character_bible(llm_client, image_client, sections, output_dir)
    generate_panels(image_client, bible, scenes, output_dir, scorer)
    page_paths = lay_out_pages(scenes, output_dir, panels_per_page)
    return export_cbz(page_paths, output_dir / "comic.cbz")


def _slugify(name: str) -> str:
    return "".join(c if c.isalnum() else "_" for c in name.lower()).strip("_")
