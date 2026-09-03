"""CLI entrypoint: python -m book_to_comic <book_path> [--output-dir DIR] [--mock]"""
from __future__ import annotations

import argparse
import sys

from book_to_comic.image_gen import MockImageClient, OpenAIImageClient
from book_to_comic.llm import AnthropicLLMClient
from book_to_comic.pipeline import run_pipeline


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Turn a book (.txt/.epub/.pdf) into a comic (.cbz).")
    parser.add_argument("book_path", help="Path to the source book file")
    parser.add_argument("--output-dir", default="comic_output", help="Where to write generated assets and the final .cbz")
    parser.add_argument("--panels-per-page", type=int, default=4)
    parser.add_argument(
        "--mock",
        action="store_true",
        help="Use offline mock LLM/image clients instead of real API calls (for testing the pipeline without API keys)",
    )
    args = parser.parse_args(argv)

    if args.mock:
        from book_to_comic.demo import FakeLLMClient

        llm_client = FakeLLMClient()
        image_client = MockImageClient()
    else:
        llm_client = AnthropicLLMClient()
        image_client = OpenAIImageClient()

    cbz_path = run_pipeline(args.book_path, args.output_dir, llm_client, image_client, panels_per_page=args.panels_per_page)
    print(f"Comic written to {cbz_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
