"""Sync the canonical manual into the Starlight content tree without changing it."""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
MANUAL = ROOT / "docs/manual"
TARGET = ROOT / "website/src/content/docs/manual"
SIDEBAR = ROOT / "website/src/data/sidebar.json"
LINK = re.compile(r"\]\((?!https?://|mailto:|#)([^)#]+)\.md(#[^)]+)?\)")
ANCHOR = re.compile(r"^<a id=\"([^\"]+)\"></a>\r?\n(##[^\r\n]+)", re.MULTILINE)


def normalize_base(raw: str) -> str:
    cleaned = raw.strip().strip("/")
    return f"/{cleaned}" if cleaned else ""


def sidebar_slugs(nodes):
    for node in nodes:
        if isinstance(node, dict) and "slug" in node:
            yield node["slug"]
        if isinstance(node, dict) and isinstance(node.get("items"), list):
            yield from sidebar_slugs(node["items"])


def read_title(text: str) -> str:
    match = re.search(r"^#\s+(.+)$", text, re.MULTILINE)
    if not match:
        raise ValueError("Manual page has no level-one heading")
    return match.group(1).strip()


def description(text: str) -> str:
    body = re.sub(r"^#\s+.+$", "", text, count=1, flags=re.MULTILINE)
    body = re.sub(r"^<a id=.*?</a>\s*$", "", body, flags=re.MULTILINE)
    body = re.sub(r"!\[[^]]*\]\([^)]*\)", "", body)
    body = re.sub(r"\[([^]]+)\]\([^)]*\)", r"\1", body)
    body = re.sub(r"[*_`>|#-]", "", body)
    paragraphs = [re.sub(r"\s+", " ", p).strip() for p in body.split("\n\n")]
    for paragraph in paragraphs:
        if paragraph and not paragraph.startswith("<"):
            return paragraph[:155].rstrip() + ("…" if len(paragraph) > 155 else "")
    return "RATools for Word 使用手册章节。"


def rewrite_links(text: str, base_manual: str) -> str:
    def replace(match: re.Match[str]) -> str:
        page = match.group(1).strip()
        fragment = match.group(2) or ""
        slug = Path(page).stem
        target = f"{base_manual}/" if slug == "index" else f"{base_manual}/{slug}/"
        return f"]({target}{fragment})"

    return LINK.sub(replace, text)


def embed_anchors(text: str) -> str:
    """Keep compatibility anchors while leaving heading-generated slugs unchanged."""
    return ANCHOR.sub(r'<a id="\1"></a>\n\n\2', text)


def rewrite_code_languages(text: str) -> str:
    """Map source code languages to names available in the site highlighter."""
    return text.replace("```vba", "```vb")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--base-path",
        default=os.environ.get("DOCS_BASE", "/RATools-for-Word"),
        help="Site base path prefix (default: $DOCS_BASE or /RATools-for-Word)",
    )
    args = parser.parse_args(argv)
    base_manual = f"{normalize_base(args.base_path)}/manual"

    pages = sorted(MANUAL.glob("*.md"))
    expected = {"manual" if p.name == "index.md" else f"manual/{p.stem}" for p in pages}
    sidebar = json.loads(SIDEBAR.read_text(encoding="utf-8"))
    configured = list(sidebar_slugs(sidebar))
    if len(configured) != len(set(configured)):
        print("Duplicate sidebar entries", file=sys.stderr)
        return 1
    if set(configured) != expected:
        print(f"Sidebar mismatch: {sorted(set(configured) ^ expected)}", file=sys.stderr)
        return 1

    TARGET.mkdir(parents=True, exist_ok=True)
    names = {p.name for p in pages}
    stale_pages = [p for p in TARGET.glob("*.md") if p.name not in names]
    if stale_pages:
        print("Generated pages are no longer in the manual; move them to the Recycle Bin:", file=sys.stderr)
        for stale in stale_pages:
            print(f"  {stale.relative_to(ROOT)}", file=sys.stderr)
        return 1

    target_assets = TARGET / "assets"
    target_assets.mkdir(exist_ok=True)
    public_assets = ROOT / "website/public/manual/assets"
    public_assets.mkdir(parents=True, exist_ok=True)
    source_assets = {p.name for p in (MANUAL / "assets").iterdir() if p.is_file()}
    existing_assets = (
        {p.name for p in target_assets.iterdir() if p.is_file()}
        | {p.name for p in public_assets.iterdir() if p.is_file()}
    )
    stale_assets = sorted(existing_assets - source_assets)
    if stale_assets:
        print("Generated assets are no longer in the manual; move them to the Recycle Bin:", file=sys.stderr)
        for stale in stale_assets:
            print(f"  {stale}", file=sys.stderr)
        return 1
    for name in sorted(source_assets):
        asset = MANUAL / "assets" / name
        (target_assets / name).write_bytes(asset.read_bytes())
        (public_assets / name).write_bytes(asset.read_bytes())

    for page in pages:
        source = page.read_text(encoding="utf-8")
        title = read_title(source)
        desc = description(source).replace('"', "'")
        body = re.sub(r"^#\s+.+\r?\n", "", source, count=1, flags=re.MULTILINE).lstrip("\r\n")
        rewritten = embed_anchors(rewrite_code_languages(rewrite_links(body, base_manual)))
        edit = f"https://github.com/PharmaRA/RATools-for-Word/edit/main/docs/manual/{page.name}"
        front = "---\n" f'title: "{title}"\n' f'description: "{desc}"\n' f'editUrl: "{edit}"\n' "---\n\n"
        (TARGET / page.name).write_text(front + rewritten, encoding="utf-8")

    print(f"Synced {len(pages)} manual pages and {len(list((MANUAL / 'assets').iterdir()))} assets")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
