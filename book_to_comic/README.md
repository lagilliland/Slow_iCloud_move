# book-to-comic

Turns a book (`.txt` / `.epub` / `.pdf`) into a paneled comic (`.cbz`), using an LLM to break the
text into scenes and panels, and an image model to draw each panel.

## Pipeline

```
book file
  -> ingest.load_text + chunk_into_sections        (raw text -> chapter-ish chunks)
  -> llm.segment_section (per chunk)                (chunk -> scenes -> panels + character mentions)
  -> llm.merge_character_descriptions               (all mentions -> one canonical "character bible" entry each)
  -> pipeline.build_character_bible                 (bible entry -> one reference portrait image, generated once)
  -> pipeline.generate_panels                       (each panel rendered, conditioned on its characters' portraits)
  -> layout.render_page                             (panels -> paneled pages with dialogue/caption boxes)
  -> pipeline.export_cbz                            (pages -> comic.cbz)
```

## Reducing character drift

Image models don't remember earlier generations, so asking for "Alex, a young man in a green
jacket" a hundred times across a book produces a hundred different-looking Alexes. Two
complementary mechanisms address this:

1. **Reference-portrait conditioning (primary).** Each character gets exactly one canonical
   text description (first mention wins — see `llm.merge_character_descriptions`) and one
   reference portrait, generated once and reused. Every later panel featuring that character is
   generated with the portrait passed in as an image input (`ImageClient.generate_with_references`,
   e.g. `images.edit` with the portrait as the reference image), not just the text description.
   This is the same idea as IP-Adapter / image-to-image reference conditioning: the model is
   shown what the character looks like instead of being asked to reimagine it from words each
   time. A fixed style suffix (`image_gen.STYLE_SUFFIX`) is appended to every prompt so the
   linework/coloring style also stays consistent panel to panel.

2. **Similarity-check retry (safety net).** After a panel is generated, `consistency.py` scores
   its similarity against the character's reference portrait(s) and, if it falls below a
   threshold, regenerates with a bumped (but still deterministic, character-seeded) seed, up to
   a retry cap. The default scorer (`PerceptualHashSimilarityScorer`) is dependency-light and
   catches gross drift; swap in `ClipEmbeddingSimilarityScorer` (needs `open-clip-torch`) for
   tighter identity-level matching when running with a real backend and a GPU. Both implement
   the same `SimilarityScorer` interface, so swapping is a one-line change in `pipeline.py`.

Neither mechanism is perfect alone — reference conditioning can still drift on pose/outfit
details, and a similarity threshold can be fooled by a wrong-but-similarly-colored panel — but
together they catch the two failure modes that matter most for a comic: "doesn't look like the
same person" and "wrong art style."

## Usage

```bash
pip install -r requirements.txt

# Offline dry run, no API keys needed (uses FakeLLMClient + MockImageClient):
python -m book_to_comic path/to/book.txt --output-dir comic_output --mock

# Real run (requires ANTHROPIC_API_KEY and OPENAI_API_KEY in the environment):
python -m book_to_comic path/to/book.epub --output-dir comic_output
```

Output layout:

```
comic_output/
  characters/<name>.png     # one reference portrait per character (the anti-drift anchor)
  panels/sceneNNN_panelNN.png
  pages/page_NNN.png
  comic.cbz
```

## Extending

- `llm.py` / `image_gen.py` define `LLMClient` / `ImageClient` ABCs — add a new backend (e.g.
  Replicate + SDXL with IP-Adapter, or a local ComfyUI pipeline) by implementing those and
  wiring it up in `cli.py`.
- `layout.py` uses a simple fixed grid; a real comic app would want varied panel sizes/shapes
  per beat (splash panels, insets) — that's the natural next step once panel-level "importance"
  is available from the LLM segmentation step.
- Speech bubbles are currently plain rectangular caption boxes; tail-pointing speech bubbles
  keyed to a speaker position would be a layout.py enhancement, not a pipeline change.

## Tests

```bash
python -m pytest tests/
```

`tests/test_pipeline_mock.py` runs the full pipeline offline. `tests/test_consistency.py` unit
tests the retry/seed-bump logic in isolation.
