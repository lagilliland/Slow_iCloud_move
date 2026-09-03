"""Load a book's raw text from .txt, .epub, or .pdf and split it into chapter-sized chunks."""
from __future__ import annotations

import re
from pathlib import Path


def load_text(path: str | Path) -> str:
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == ".txt":
        return path.read_text(encoding="utf-8", errors="ignore")
    if suffix == ".epub":
        return _load_epub(path)
    if suffix == ".pdf":
        return _load_pdf(path)
    raise ValueError(f"Unsupported book format: {suffix}")


def _load_epub(path: Path) -> str:
    import ebooklib
    from bs4 import BeautifulSoup
    from ebooklib import epub

    book = epub.read_epub(str(path))
    parts = []
    for item in book.get_items():
        if item.get_type() == ebooklib.ITEM_DOCUMENT:
            soup = BeautifulSoup(item.get_content(), "html.parser")
            text = soup.get_text(separator="\n").strip()
            if text:
                parts.append(text)
    return "\n\n".join(parts)


def _load_pdf(path: Path) -> str:
    from pypdf import PdfReader

    reader = PdfReader(str(path))
    return "\n\n".join(page.extract_text() or "" for page in reader.pages)


def chunk_into_sections(text: str, max_chars: int = 6000) -> list[str]:
    """Split text into LLM-sized chunks, preferring chapter/paragraph boundaries."""
    chapter_pattern = re.compile(r"\n\s*(chapter\s+\w+|CHAPTER\s+\w+)\b", re.IGNORECASE)
    boundaries = [m.start() for m in chapter_pattern.finditer(text)]

    if len(boundaries) >= 2:
        chapters = []
        for i, start in enumerate(boundaries):
            end = boundaries[i + 1] if i + 1 < len(boundaries) else len(text)
            chapters.append(text[start:end].strip())
    else:
        chapters = [text]

    sections: list[str] = []
    for chapter in chapters:
        if len(chapter) <= max_chars:
            sections.append(chapter)
            continue
        paragraphs = chapter.split("\n\n")
        current = ""
        for para in paragraphs:
            if len(current) + len(para) + 2 > max_chars and current:
                sections.append(current.strip())
                current = ""
            current += para + "\n\n"
        if current.strip():
            sections.append(current.strip())
    return [s for s in sections if s.strip()]
