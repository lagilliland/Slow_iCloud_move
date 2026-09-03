"""Compose generated panel images into comic pages with borders, captions, and speech bubbles."""
from __future__ import annotations

from pathlib import Path

from book_to_comic.models import ComicPage, Panel

PAGE_SIZE = (1600, 2400)
MARGIN = 24
GUTTER = 16


def _grid_for(count: int) -> tuple[int, int]:
    """Rows, cols for a simple panel grid; good enough for an MVP layout."""
    if count <= 1:
        return 1, 1
    if count <= 2:
        return 2, 1
    if count <= 4:
        return 2, 2
    if count <= 6:
        return 3, 2
    return 4, 2


def render_page(panels: list[Panel], page_index: int, output_dir: str | Path) -> ComicPage:
    from PIL import Image, ImageDraw, ImageFont

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    page = Image.new("RGB", PAGE_SIZE, color="white")
    draw = ImageDraw.Draw(page)
    try:
        font = ImageFont.truetype("DejaVuSans.ttf", 22)
    except OSError:
        font = ImageFont.load_default()

    rows, cols = _grid_for(len(panels))
    cell_w = (PAGE_SIZE[0] - 2 * MARGIN - (cols - 1) * GUTTER) // cols
    cell_h = (PAGE_SIZE[1] - 2 * MARGIN - (rows - 1) * GUTTER) // rows

    for i, panel in enumerate(panels):
        if panel.image_path is None:
            continue
        row, col = divmod(i, cols)
        x = MARGIN + col * (cell_w + GUTTER)
        y = MARGIN + row * (cell_h + GUTTER)

        panel_img = Image.open(panel.image_path).convert("RGB")
        panel_img = _fit_cover(panel_img, (cell_w, cell_h))
        page.paste(panel_img, (x, y))
        draw.rectangle([x, y, x + cell_w, y + cell_h], outline="black", width=4)

        text = panel.dialogue or panel.caption
        if text:
            _draw_caption(draw, text, x, y, cell_w, font)

    out_path = output_dir / f"page_{page_index:03d}.png"
    page.save(out_path)
    return ComicPage(index=page_index, panel_image_paths=[p.image_path for p in panels if p.image_path], output_path=str(out_path))


def _fit_cover(image, box: tuple[int, int]):
    from PIL import Image

    box_w, box_h = box
    src_ratio = image.width / image.height
    box_ratio = box_w / box_h
    if src_ratio > box_ratio:
        new_h = box_h
        new_w = int(new_h * src_ratio)
    else:
        new_w = box_w
        new_h = int(new_w / src_ratio)
    image = image.resize((new_w, new_h), Image.LANCZOS)
    left = (new_w - box_w) // 2
    top = (new_h - box_h) // 2
    return image.crop((left, top, left + box_w, top + box_h))


def _draw_caption(draw, text: str, x: int, y: int, width: int, font) -> None:
    import textwrap

    wrapped = textwrap.fill(text, width=30)
    lines = wrapped.split("\n")
    line_height = 26
    box_h = line_height * len(lines) + 16
    box_y = y + 8
    draw.rectangle([x + 8, box_y, x + width - 8, box_y + box_h], fill="white", outline="black", width=2)
    for i, line in enumerate(lines):
        draw.text((x + 16, box_y + 8 + i * line_height), line, fill="black", font=font)


def paginate(panels: list[Panel], panels_per_page: int = 4) -> list[list[Panel]]:
    return [panels[i : i + panels_per_page] for i in range(0, len(panels), panels_per_page)]
