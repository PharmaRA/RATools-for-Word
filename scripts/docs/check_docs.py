"""Check documented entry points and local links; never modify source or output."""

from __future__ import annotations

import argparse
import json
import re
import unicodedata
import xml.etree.ElementTree as ET
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import unquote, urlsplit

import markdown
import yaml

ROOT = Path(__file__).resolve().parents[2]
MANUAL = ROOT / "docs/manual"


class Page(HTMLParser):
    def __init__(self, source: str):
        super().__init__()
        self.ids: set[str] = set()
        self.links: list[str] = []
        self.feed(source)

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        for key in ("id", "name"):
            if attrs.get(key):
                self.ids.add(attrs[key])
        for key in ("href", "src"):
            if attrs.get(key):
                self.links.append(attrs[key])


def render(path: Path) -> Page:
    source = path.read_text(encoding="utf-8")
    if path.suffix == ".md":
        source = markdown.markdown(
            source, extensions=["toc", "tables", "attr_list", "fenced_code"]
        )
    return Page(source)


def normalize(value: str) -> str:
    return re.sub(r"[\W_]", "", unicodedata.normalize("NFKC", value)).lower()


def entry_points() -> dict[str, dict[str, tuple[str, str]]]:
    ribbon = {}
    tree = ET.parse(ROOT / "dotm/customUI/customUI14.xml")
    for item in tree.iter():
        attrs = item.attrib
        if attrs.get("onAction") or attrs.get("idMso"):
            key = "id:" + attrs["id"] if "id" in attrs else "mso:" + attrs["idMso"]
            if key in ribbon:
                raise ValueError(f"Duplicate Ribbon entry: {key}")
            ribbon[key] = (attrs.get("label", "").strip(), attrs.get("onAction", attrs.get("idMso", "")))
    registry = (ROOT / "modules/UI_MacroRegistry.bas").read_text(encoding="gbk")
    active = "\n".join(line for line in registry.splitlines() if not line.lstrip().startswith("'"))
    macros = {
        name: (label, name)
        for name, label in re.findall(r'items\.Add\s+Array\("([^"]+)"\s*,\s*_?\s*"([^"]+)"', active)
    }
    form = (ROOT / "userforms/frmQuickToolbar.frm").read_text(encoding="gbk")
    toolbar = {name: ("", name) for name in re.findall(r'RunQuickToolbarAction "([^"]+)"', form)}
    return {"ribbon": ribbon, "macros": macros, "toolbar": toolbar}


def check_coverage(errors: list[str]) -> None:
    mapping = json.loads((ROOT / "docs/feature-map.json").read_text(encoding="utf-8"))
    for category, actual in entry_points().items():
        documented = mapping[category]
        for key in sorted(actual.keys() - documented.keys()):
            errors.append(f"Undocumented {category} entry: {key}")
        for key in sorted(documented.keys() - actual.keys()):
            errors.append(f"Stale {category} entry: {key}")
        for key in actual.keys() & documented.keys():
            label, action = actual[key]
            entry = documented[key]
            if entry["action"] != action or entry["label"] != label:
                errors.append(f"Changed {category} entry needs review: {key}")
            target = urlsplit(entry["target"])
            path = (MANUAL / unquote(target.path)).resolve()
            if not path.is_relative_to(MANUAL.resolve()) or not path.is_file():
                errors.append(f"Missing manual target for {key}: {entry['target']}")
                continue
            if target.fragment and unquote(target.fragment) not in render(path).ids:
                errors.append(f"Missing manual anchor for {key}: {entry['target']}")
            if label and normalize(label) not in normalize(path.read_text(encoding="utf-8")):
                errors.append(f"Manual does not mention source label {label}: {entry['target']}")
        print(f"Coverage {category}: {len(actual)} source entries, {len(documented)} mapped")


def check_links(paths: list[Path], boundary: Path, errors: list[str], site=False, base_path="") -> int:
    cache: dict[Path, Page] = {}
    checked = 0
    boundary = boundary.resolve()
    for source in paths:
        source = source.resolve()
        page = cache.setdefault(source, render(source))
        for link in page.links:
            parts = urlsplit(link)
            if parts.scheme or parts.netloc:
                continue
            href = unquote(parts.path)
            if href.startswith("/"):
                if site and base_path and href.startswith(base_path + "/"):
                    href = href[len(base_path) + 1:]
                target = boundary / href.lstrip("/")
            else:
                target = source.parent / href if href else source
            target = target.resolve()
            if not target.is_relative_to(boundary):
                errors.append(f"Link escapes checked root: {source.name} -> {link}")
                continue
            if site and target.is_dir():
                target = target / "index.html"
            checked += 1
            if not target.exists():
                errors.append(f"Broken link: {source.relative_to(boundary)} -> {link}")
                continue
            if parts.fragment and target.suffix in (".md", ".html"):
                target_page = cache.setdefault(target, render(target))
                if unquote(parts.fragment) not in target_page.ids:
                    errors.append(f"Broken anchor: {source.relative_to(boundary)} -> {link}")
    return checked


def nav_paths(nodes):
    for node in nodes:
        if isinstance(node, dict) and "slug" in node:
            yield node["slug"]
        if isinstance(node, dict) and isinstance(node.get("items"), list):
            yield from nav_paths(node["items"])


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--site-dir", type=Path)
    parser.add_argument("--base-path", default="")
    args = parser.parse_args()
    errors: list[str] = []
    check_coverage(errors)
    pages = sorted(MANUAL.glob("*.md"))
    config = json.loads((ROOT / "website/src/data/sidebar.json").read_text(encoding="utf-8"))
    nav = list(nav_paths(config))
    if len(nav) != len(set(nav)):
        errors.append("Duplicate navigation entries")
    names = {"manual" if p.name == "index.md" else f"manual/{p.stem}" for p in pages}
    if set(nav) != names:
        errors.append(f"Navigation mismatch: {sorted(set(nav) ^ names)}")
    # Bound reads to the manual and its entry points; no repository-wide scan.
    source_pages = pages + [ROOT / "README.md", ROOT / "docs/README.md", ROOT / "docs/docs-site.md"]
    checked = check_links(source_pages, ROOT, errors)
    print(f"Markdown: {len(pages)} manual pages, {checked} local links checked")
    if args.site_dir:
        site = args.site_dir.resolve()
        if not (site / "index.html").is_file():
            errors.append(f"Site entry point missing: {site}")
        else:
            html_pages = sorted(site.rglob("*.html"))
            checked = check_links(html_pages, site, errors, site=True, base_path=args.base_path)
            if (site / "search/search_index.json").is_file():
                pass
            elif not (site / "pagefind/pagefind.js").is_file():
                errors.append("Search index missing")
            print(f"Rendered site: {len(html_pages)} HTML pages, {checked} local links checked")
    for error in errors:
        print("ERROR:", error)
    print(f"Documentation checks: {'FAIL' if errors else 'PASS'} ({len(errors)} errors)")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())