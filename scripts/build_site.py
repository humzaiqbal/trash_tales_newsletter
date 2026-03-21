#!/usr/bin/env python3
from __future__ import annotations

import datetime as dt
import html
import json
import os
import re
import shutil
import zipfile
from pathlib import Path
from typing import Dict, List, Tuple
from xml.etree import ElementTree as ET


ROOT_DIR = Path(__file__).resolve().parent.parent
SOURCE_DIR = Path(os.environ.get("TRASH_TALES_SOURCE_DIR", str(Path.home() / "Downloads" / "newsletter")))
SITE_DIR = ROOT_DIR / "site"
POSTS_DIR = SITE_DIR / "posts"
ASSETS_DIR = SITE_DIR / "assets"
IMAGES_DIR = ASSETS_DIR / "images"
CHARACTER_LIST_FILE = SOURCE_DIR / "Character List.docx"
ASSET_VERSION = "20260321b"

QUOTED_NICKNAMES = re.compile(r"[\"“”]([^\"“”]+)[\"“”]")
EPISODE_NUMBER = re.compile(r"episode_(\d+)", re.IGNORECASE)
EPISODE_LABEL = re.compile(r"episode_([0-9]+(?:-[0-9]+)?)", re.IGNORECASE)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"w": W_NS, "a": A_NS, "r": R_NS, "pr": PKG_REL_NS}

SKIP_LINES = {
    "character list",
    "i made a character list detailing every character here. if you are lost as to who someone is look here :d",
    "cool things about people",
    "heres a list of some of the best responses i’ve gotten to the question “what’s something cool or interesting about you that people wouldn’t expect?”",
    "heres a list of some of the best responses i’ve gotten to the question \"what’s something cool or interesting about you that people wouldn’t expect?\"",
}

SPECIAL_TITLES = {
    "62": "Episode 62- Congrats to my CFAs!!",
    "64": "Episode 64 - New website",
}

EXCERPT_SKIP_LINES = SKIP_LINES | {
    "previous episodes",
    "if you want to find previous episodes look here.",
    "if you want to see any of the previous newsletters look here.",
}

EPISODE_TEXT_REPLACEMENTS = {
    "65": {
        "Tightly Knit": "Tight Knit",
        "The Spikeballer": "The Spiker",
        "The spikeballer": "The Spiker",
    }
}

INLINE_LINKS = [
    (
        "luck surface area",
        "https://www.codusoperandi.com/posts/increasing-your-luck-surface-area",
    ),
    (
        "Assessor Recorder",
        "https://www.sf.gov/departments--assessor-recorder",
    ),
    (
        "a Yelp Review",
        "https://www.yelp.com/biz/kowloon-tong-dessert-cafe-san-francisco?hrid=oyF7m2y0KoziaZhPGojM6A&utm_campaign=www_review_share_popup&utm_medium=copy_link&utm_source=(direct)",
    ),
]


def canonical_alias(alias: str) -> str:
    return alias.strip().strip('"').strip("“").strip("”").strip()


def normalize_alias_lookup(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", text.lower())


def is_skippable_line(line: str) -> bool:
    low = line.strip().lower()
    return low in SKIP_LINES


def is_excerpt_skippable_line(line: str) -> bool:
    low = line.strip().lower()
    return low in EXCERPT_SKIP_LINES


def normalize_episode_text(text: str, episode_label: str) -> str:
    replacements = EPISODE_TEXT_REPLACEMENTS.get(episode_label, {})
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text


def render_inline_html(
    line: str,
    variant_lookup: Dict[str, Tuple[str, str]],
    variant_pattern: re.Pattern,
) -> str:
    remaining = line
    rendered: List[str] = []

    while remaining:
        earliest_match = None
        earliest_phrase = None
        earliest_url = None

        for phrase, url in INLINE_LINKS:
            pattern = re.compile(rf"{re.escape(phrase)}\s+{re.escape(url)}")
            match = pattern.search(remaining)
            if not match:
                continue
            if earliest_match is None or match.start() < earliest_match.start():
                earliest_match = match
                earliest_phrase = phrase
                earliest_url = url

        if earliest_match is None:
            rendered.append(inject_character_tooltips(remaining, variant_lookup, variant_pattern))
            break

        before = remaining[:earliest_match.start()]
        rendered.append(inject_character_tooltips(before, variant_lookup, variant_pattern))
        rendered.append(
            f'<a href="{html.escape(earliest_url)}" target="_blank" rel="noreferrer">{html.escape(earliest_phrase)}</a>'
        )
        remaining = remaining[earliest_match.end():]

    return "".join(rendered)


def read_docx_blocks(
    docx_path: Path,
    image_output_dir: Path | None = None,
    image_url_prefix: str | None = None,
) -> List[dict]:
    blocks: List[dict] = []
    with zipfile.ZipFile(docx_path) as zf:
        doc_xml = zf.read("word/document.xml")
        rels_map: Dict[str, str] = {}
        if "word/_rels/document.xml.rels" in zf.namelist():
            rels_root = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
            for rel in rels_root.findall("pr:Relationship", NS):
                rid = rel.attrib.get("Id")
                target = rel.attrib.get("Target", "")
                if rid and target:
                    rels_map[rid] = target

        root = ET.fromstring(doc_xml)
        image_count = 0

        for para in root.findall(".//w:p", NS):
            text_parts: List[str] = []
            para_images: List[str] = []
            for node in para.iter():
                if node.tag == f"{{{W_NS}}}t":
                    text_parts.append(node.text or "")
                elif node.tag == f"{{{A_NS}}}blip":
                    rid = node.attrib.get(f"{{{R_NS}}}embed")
                    if not rid:
                        continue
                    target = rels_map.get(rid)
                    if not target or not image_output_dir or not image_url_prefix:
                        continue
                    internal_path = "word/" + target.lstrip("/")
                    if internal_path not in zf.namelist():
                        continue
                    image_count += 1
                    ext = Path(target).suffix.lower() or ".png"
                    filename = f"img-{image_count:03d}{ext}"
                    out_path = image_output_dir / filename
                    out_path.write_bytes(zf.read(internal_path))
                    para_images.append(f"{image_url_prefix}/{filename}")

            text = "".join(text_parts).strip()
            if text:
                blocks.append({"type": "paragraph", "text": text})
            for image_url in para_images:
                blocks.append({"type": "image", "url": image_url})
    return blocks


def read_docx_paragraph_lines(docx_path: Path) -> List[str]:
    blocks = read_docx_blocks(docx_path)
    return [b["text"] for b in blocks if b.get("type") == "paragraph" and b.get("text")]


def parse_character_list(lines: List[str]) -> Dict[str, str]:
    aliases: Dict[str, str] = {}
    for line in lines:
        q_matches = list(QUOTED_NICKNAMES.finditer(line))
        if not q_matches:
            continue
        if not (line.strip().startswith("“") or line.strip().startswith('"')):
            continue

        # Only aliases at the beginning of the line are alias declarations.
        # Later quoted names can appear in the description itself.
        leading_matches = []
        idx = 0
        while idx < len(line) and line[idx].isspace():
            idx += 1
        for m in q_matches:
            if m.start() != idx:
                break
            leading_matches.append(m)
            idx = m.end()
            while idx < len(line) and line[idx].isspace():
                idx += 1

        if not leading_matches:
            continue

        desc = line[leading_matches[-1].end() :].strip(" :-.\t")
        if not desc:
            continue

        for m in leading_matches:
            key = canonical_alias(m.group(1).strip())
            if key and key not in aliases:
                aliases[key] = desc
    return aliases


def expand_alias_candidates(alias: str) -> set[str]:
    candidates = {alias}
    # Allow both "YT" and "Youtube" spellings to map to same character.
    candidates.add(re.sub(r"\byt\b", "youtube", alias, flags=re.IGNORECASE))
    candidates.add(re.sub(r"\byoutube\b", "yt", alias, flags=re.IGNORECASE))
    return {c.strip() for c in candidates if c.strip()}


def build_variant_lookup(character_defs: Dict[str, str]) -> Tuple[Dict[str, Tuple[str, str]], re.Pattern]:
    alias_lookup: Dict[str, Tuple[str, str]] = {}
    for alias, desc in character_defs.items():
        for candidate in expand_alias_candidates(alias):
            norm = normalize_alias_lookup(candidate)
            alias_lookup[norm] = (alias, desc)
            if not candidate.lower().startswith("the "):
                alias_lookup[normalize_alias_lookup(f"the {candidate}")] = (alias, desc)
    pattern = re.compile(r"([“\"])\s*([^\"“”]+?)((?:[’']s)?[,.!?]?)\s*([”\"])", re.IGNORECASE)
    return alias_lookup, pattern


def inject_character_tooltips(
    line: str,
    variant_lookup: Dict[str, Tuple[str, str]],
    variant_pattern: re.Pattern,
) -> str:
    if not line:
        return ""
    if not variant_lookup:
        return html.escape(line)

    out: List[str] = []
    last_idx = 0
    for m in variant_pattern.finditer(line):
        start, end = m.span()
        out.append(html.escape(line[last_idx:start]))
        open_q, alias_txt, suffix, close_q = m.groups()
        alias, desc = variant_lookup.get(normalize_alias_lookup(alias_txt), ("", ""))
        if not alias:
            out.append(html.escape(line[start:end]))
        else:
            visible = f"{open_q}{alias_txt}{suffix}{close_q}"
            out.append(
                f'<span class="character-chip" tabindex="0" '
                f'data-character="{html.escape(alias)}" '
                f'data-description="{html.escape(desc)}">{html.escape(visible)}</span>'
            )
        last_idx = end
    out.append(html.escape(line[last_idx:]))
    return "".join(out)


def block_to_html(block: dict, variant_lookup: Dict[str, Tuple[str, str]], variant_pattern: re.Pattern) -> str:
    if block["type"] == "image":
        return (
            '<figure class="post-image-wrap">'
            f'<img class="post-image" src="{html.escape(block["url"])}" loading="lazy" alt="Newsletter image" />'
            "</figure>"
        )
    line = block["text"].strip()
    if not line or is_skippable_line(line):
        return ""
    lowered = line.lower()
    weekdays = {"monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"}
    if lowered in weekdays:
        return f"<h2>{html.escape(line)}</h2>"
    return f"<p>{render_inline_html(line, variant_lookup, variant_pattern)}</p>"


def parse_episode_number(name: str) -> int:
    m = EPISODE_NUMBER.search(name)
    return int(m.group(1)) if m else -1


def parse_episode_label(name: str) -> str:
    m = EPISODE_LABEL.search(name)
    return m.group(1) if m else "unknown"


def render_episode_sidebar(
    posts: List[dict],
    current_url: str | None = None,
    url_prefix: str = "",
) -> str:
    items = []
    for post in posts:
        href = post["url"]
        if url_prefix and href.startswith("./"):
            href = f"{url_prefix}{href[2:]}"
        elif url_prefix:
            href = f"{url_prefix}{href}"
        is_current = post["url"] == current_url
        active_class = " active" if is_current else ""
        current_attr = ' aria-current="page"' if is_current else ""
        items.append(
            f'<li><a class="episode-link{active_class}" href="{html.escape(href)}"{current_attr}>'
            f"{html.escape(post['title'])}</a></li>"
        )
    items_html = "".join(items)
    return f"""
    <aside class="episode-sidebar" aria-label="Episode navigation">
      <div class="episode-sidebar-inner episode-sidebar-desktop">
        <p class="episode-sidebar-label">Browse</p>
        <h2 class="episode-sidebar-title">Episodes</h2>
        <p class="episode-sidebar-meta">{len(posts)} episodes</p>
        <ol class="episode-list">
          {items_html}
        </ol>
      </div>
      <details class="episode-sidebar-mobile">
        <summary class="episode-sidebar-toggle">
          <span>Episodes</span>
          <span class="episode-sidebar-toggle-meta">{len(posts)} episodes</span>
        </summary>
        <div class="episode-sidebar-mobile-body">
          <ol class="episode-list">
            {items_html}
          </ol>
        </div>
      </details>
    </aside>
"""


def render_post_html(
    title: str,
    article_html: str,
    post_date: str,
    canonical_name: str,
    posts: List[dict],
    current_url: str,
) -> str:
    sidebar_html = render_episode_sidebar(posts, current_url=current_url, url_prefix="../")
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{html.escape(title)} | Trash Tales</title>
  <meta name="description" content="Weekly Trash Tales newsletter archive by Trash Tales." />
  <link rel="stylesheet" href="../assets/styles.css?v={ASSET_VERSION}" />
</head>
<body>
  <header class="site-header">
    <div class="site-header-inner">
      <a class="brand" href="../index.html">TRASH TALES</a>
    </div>
  </header>

  <main class="container page-shell">
    {sidebar_html}
    <article class="post page-main">
      <p class="post-meta">{post_date}</p>
      <h1>{html.escape(title)}</h1>
      <section class="post-content">
        {article_html}
      </section>
      <footer class="post-footer">
        <p>Source file: <code>{html.escape(canonical_name)}</code></p>
      </footer>
    </article>
  </main>

  <div id="tooltip" class="tooltip" role="status" aria-live="polite"></div>
  <script src="../assets/app.js?v={ASSET_VERSION}"></script>
</body>
</html>
"""


def render_post_card(post: dict, summary: str) -> str:
    return f"""
      <article class="post-card">
        <p class="post-meta">{html.escape(post['date'])}</p>
        <h2><a href="{html.escape(post['url'])}">{html.escape(post['title'])}</a></h2>
        <p>{html.escape(summary)}</p>
        <a class="read-more" href="{html.escape(post['url'])}">Read post</a>
      </article>
"""


def render_index_html(posts: List[dict]) -> str:
    cards_html = "".join(render_post_card(p, p["excerpt"]) for p in posts)
    sidebar_html = render_episode_sidebar(posts)
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Trash Tales Newsletter</title>
  <meta name="description" content="Weekly Trash Tales newsletter archive by Trash Tales." />
  <link rel="stylesheet" href="./assets/styles.css?v={ASSET_VERSION}" />
</head>
<body>
  <header class="site-header">
    <div class="site-header-inner">
      <a class="brand" href="./index.html">TRASH TALES</a>
    </div>
  </header>

  <main class="container page-shell">
    {sidebar_html}
    <div class="page-main">
      <section class="hero">
        <h1>Weekly Newsletter Archive</h1>
        <p>
          Notes, stories, and characters from each week. Hover or tap a character name to see who they are.
        </p>
      </section>
      <section class="search-panel">
        <label class="search-label" for="archive-search">Search newsletters</label>
        <input
          id="archive-search"
          class="search-input"
          type="search"
          placeholder="Search keywords across all episodes"
          autocomplete="off"
          spellcheck="false"
        />
        <p class="search-help">Search titles and full article text across the archive.</p>
        <p id="search-status" class="search-status" hidden></p>
      </section>
      <section id="archive-grid" class="archive-grid">
        {cards_html}
      </section>
    </div>
  </main>
  <script src="./assets/app.js?v={ASSET_VERSION}" data-search-index="./assets/search-index.json?v={ASSET_VERSION}"></script>
</body>
</html>
"""


def write_assets() -> None:
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)

    styles = """* {
  box-sizing: border-box;
}

:root {
  --bg: #fcfcfa;
  --text: #1f2328;
  --muted: #59636e;
  --line: #d7dde4;
  --accent: #0b57d0;
  --chip: #eef3ff;
}

html, body {
  margin: 0;
  padding: 0;
  background: var(--bg);
  color: var(--text);
  font-family: "Inter", "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  line-height: 1.7;
}

.site-header {
  border-bottom: 1px solid var(--line);
  position: sticky;
  top: 0;
  background: color-mix(in oklab, var(--bg) 92%, white 8%);
  backdrop-filter: blur(4px);
  z-index: 20;
}

.site-header-inner {
  max-width: 1180px;
  margin: 0 auto;
  padding: 14px 20px;
  display: flex;
  align-items: center;
}

.brand {
  color: var(--text);
  text-decoration: none;
  font-weight: 700;
  letter-spacing: 0.08em;
  font-size: 0.9rem;
}

nav a {
  color: var(--muted);
  text-decoration: none;
  font-size: 0.95rem;
}

.container {
  max-width: 1180px;
  margin: 0 auto;
  padding: 30px 20px 64px;
}

.page-shell {
  display: grid;
  grid-template-columns: minmax(170px, 220px) minmax(0, 1fr);
  gap: 24px;
  align-items: start;
}

.page-main {
  min-width: 0;
}

.episode-sidebar {
  align-self: start;
}

.episode-sidebar-inner {
  position: sticky;
  top: 84px;
  max-height: calc(100vh - 104px);
  overflow: auto;
  border: 1px solid var(--line);
  border-radius: 12px;
  padding: 16px;
  background: color-mix(in oklab, white 72%, var(--bg) 28%);
}

.episode-sidebar-mobile {
  display: none;
  border: 1px solid var(--line);
  border-radius: 12px;
  background: color-mix(in oklab, white 72%, var(--bg) 28%);
}

.episode-sidebar-toggle {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 14px 16px;
  cursor: pointer;
  font-weight: 600;
  list-style: none;
}

.episode-sidebar-toggle::-webkit-details-marker {
  display: none;
}

.episode-sidebar-toggle::after {
  content: "+";
  color: var(--muted);
  font-size: 1.1rem;
  line-height: 1;
}

.episode-sidebar-mobile[open] .episode-sidebar-toggle::after {
  content: "-";
}

.episode-sidebar-toggle-meta {
  margin-left: auto;
  padding-right: 8px;
  color: var(--muted);
  font-weight: 500;
  font-size: 0.92rem;
}

.episode-sidebar-mobile-body {
  padding: 0 12px 12px;
}

.episode-sidebar-label {
  margin: 0 0 2px;
  color: var(--muted);
  font-size: 0.78rem;
  font-weight: 700;
  letter-spacing: 0.08em;
  text-transform: uppercase;
}

.episode-sidebar-title {
  margin: 0;
  font-size: 1.05rem;
}

.episode-sidebar-meta {
  margin: 2px 0 14px;
  color: var(--muted);
  font-size: 0.92rem;
}

.episode-list {
  list-style: none;
  margin: 0;
  padding: 0;
  display: grid;
  gap: 6px;
}

.episode-link {
  display: block;
  padding: 8px 10px;
  border-radius: 8px;
  color: var(--muted);
  text-decoration: none;
  overflow-wrap: anywhere;
  transition: background 120ms ease, color 120ms ease;
}

.episode-link:hover,
.episode-link:focus-visible {
  background: #fff;
  color: var(--text);
  outline: none;
}

.episode-link.active {
  background: #fff;
  color: var(--text);
  font-weight: 600;
  box-shadow: inset 0 0 0 1px var(--line);
}

.hero h1,
.post h1 {
  font-family: "Georgia", "Times New Roman", serif;
  font-weight: 700;
  line-height: 1.2;
  letter-spacing: -0.01em;
}

.hero h1 {
  margin: 0 0 10px;
  font-size: clamp(1.9rem, 3.2vw, 2.8rem);
}

.hero p {
  margin: 0 0 24px;
  color: var(--muted);
  max-width: 680px;
}

.search-panel {
  margin: 0 0 24px;
}

.search-label {
  display: block;
  margin: 0 0 8px;
  font-weight: 600;
  color: var(--text);
}

.search-input {
  width: 100%;
  padding: 13px 14px;
  border: 1px solid var(--line);
  border-radius: 10px;
  background: #fff;
  color: var(--text);
  font: inherit;
}

.search-input:focus {
  outline: 2px solid color-mix(in oklab, var(--accent) 30%, white 70%);
  border-color: var(--accent);
}

.search-help,
.search-status {
  margin: 8px 0 0;
  color: var(--muted);
  font-size: 0.95rem;
}

.archive-grid {
  display: grid;
  grid-template-columns: 1fr;
  gap: 14px;
}

.post-card {
  border: 1px solid var(--line);
  border-radius: 10px;
  padding: 16px;
  background: #fff;
}

.post-card h2 {
  margin: 0 0 8px;
  font-family: "Georgia", "Times New Roman", serif;
  font-size: clamp(1.25rem, 2.2vw, 1.6rem);
}

.post-card h2 a {
  color: var(--text);
  text-decoration: none;
}

.post-card p {
  margin: 0 0 10px;
  color: var(--muted);
}

.post-meta {
  color: var(--muted);
  font-size: 0.88rem;
  letter-spacing: 0.01em;
  margin: 0 0 10px;
}

.read-more {
  color: var(--accent);
  text-decoration: none;
  font-weight: 600;
}

.post h1 {
  margin: 0 0 6px;
  font-size: clamp(2rem, 3.6vw, 3rem);
}

.post-content {
  font-family: "Georgia", "Times New Roman", serif;
  font-size: clamp(1.06rem, 1.4vw, 1.18rem);
}

.post-content p,
.post-content h2,
.post-content h3 {
  max-width: 74ch;
}

.post-content p {
  margin: 0 0 1.15em;
}

.post-content h2 {
  margin: 1.8em 0 0.65em;
  font-size: 1.25em;
}

.post-image-wrap {
  margin: 1.2em 0 1.5em;
}

.post-image {
  display: block;
  width: 100%;
  max-width: 740px;
  height: auto;
  border-radius: 10px;
  border: 1px solid var(--line);
  background: #fff;
}

.post-footer {
  margin-top: 28px;
  padding-top: 10px;
  border-top: 1px solid var(--line);
  color: var(--muted);
  font-size: 0.9rem;
}

.character-chip {
  background: var(--chip);
  border-bottom: 1px dashed var(--accent);
  border-radius: 4px;
  padding: 0 2px;
  cursor: help;
  transition: background 120ms ease;
}

.character-chip:hover,
.character-chip:focus-visible,
.character-chip.active {
  background: #dae6ff;
  outline: none;
}

.tooltip {
  position: fixed;
  z-index: 30;
  max-width: min(320px, calc(100vw - 24px));
  background: #121821;
  color: #f5f8ff;
  border-radius: 8px;
  padding: 10px 12px;
  box-shadow: 0 12px 30px rgba(0, 0, 0, 0.3);
  font-size: 0.9rem;
  line-height: 1.45;
  pointer-events: none;
  opacity: 0;
  transform: translateY(4px);
  transition: opacity 120ms ease, transform 120ms ease;
}

.tooltip.show {
  opacity: 1;
  transform: translateY(0);
}

@media (max-width: 760px) {
  .page-shell {
    grid-template-columns: minmax(150px, 190px) minmax(0, 1fr);
    gap: 16px;
  }

  .episode-sidebar-inner {
    padding: 12px;
  }

  .episode-link {
    padding: 7px 8px;
    font-size: 0.95rem;
  }
}

@media (max-width: 560px) {
  .page-shell {
    grid-template-columns: 1fr;
    gap: 20px;
  }

  .episode-sidebar-desktop {
    display: none;
  }

  .episode-sidebar-mobile {
    display: block;
  }

  .episode-sidebar-mobile .episode-list {
    max-height: min(55vh, 360px);
    overflow: auto;
    padding-right: 2px;
  }

  .episode-sidebar-mobile .episode-link {
    background: #fff;
    box-shadow: inset 0 0 0 1px var(--line);
  }
}

@media (max-width: 640px) {
  .site-header-inner,
  .container {
    padding-left: 14px;
    padding-right: 14px;
  }

  .post-card {
    border-radius: 8px;
    padding: 14px;
  }

  .tooltip {
    font-size: 0.86rem;
  }
}
"""
    (ASSETS_DIR / "styles.css").write_text(styles, encoding="utf-8")

    app_js = """(() => {
  const tooltip = document.getElementById("tooltip");
  if (!tooltip) return;

  let active = null;

  function positionTip(clientX, clientY) {
    const pad = 12;
    const rect = tooltip.getBoundingClientRect();
    let left = clientX + 12;
    let top = clientY + 14;

    if (left + rect.width > window.innerWidth - pad) {
      left = window.innerWidth - rect.width - pad;
    }
    if (top + rect.height > window.innerHeight - pad) {
      top = clientY - rect.height - 14;
    }
    if (left < pad) left = pad;
    if (top < pad) top = pad;

    tooltip.style.left = left + "px";
    tooltip.style.top = top + "px";
  }

  function showTip(el, x, y) {
    const character = el.dataset.character || "Character";
    const description = el.dataset.description || "";
    tooltip.replaceChildren();
    const strong = document.createElement("strong");
    strong.textContent = character;
    const br = document.createElement("br");
    const text = document.createTextNode(description);
    tooltip.append(strong, br, text);
    tooltip.classList.add("show");
    el.classList.add("active");
    active = el;
    positionTip(x, y);
  }

  function hideTip() {
    tooltip.classList.remove("show");
    if (active) active.classList.remove("active");
    active = null;
  }

  function wire(el) {
    el.addEventListener("mouseenter", (e) => showTip(el, e.clientX, e.clientY));
    el.addEventListener("mousemove", (e) => positionTip(e.clientX, e.clientY));
    el.addEventListener("mouseleave", hideTip);

    el.addEventListener("focus", () => {
      const r = el.getBoundingClientRect();
      showTip(el, r.left + r.width / 2, r.top + r.height / 2);
    });
    el.addEventListener("blur", hideTip);

    el.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const r = el.getBoundingClientRect();
      if (active === el) {
        hideTip();
      } else {
        showTip(el, r.left + r.width / 2, r.top + r.height / 2);
      }
    });
  }

  document.querySelectorAll(".character-chip").forEach(wire);
  document.addEventListener("click", (e) => {
    if (!(e.target instanceof Element)) return;
    if (!e.target.closest(".character-chip")) hideTip();
  });
  window.addEventListener("scroll", () => {
    if (active) {
      const r = active.getBoundingClientRect();
      positionTip(r.left + r.width / 2, r.top + r.height / 2);
    }
  }, { passive: true });
})();

(() => {
  const input = document.getElementById("archive-search");
  const grid = document.getElementById("archive-grid");
  const status = document.getElementById("search-status");
  if (!input || !grid || !status) return;

  const scriptEl = document.querySelector("script[data-search-index]");
  const searchIndexUrl = scriptEl?.dataset.searchIndex;
  if (!searchIndexUrl) return;

  const initialHtml = grid.innerHTML;
  let posts = [];
  let loaded = false;

  function escapeHtml(value) {
    return value
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  function normalize(value) {
    return value.toLowerCase().replace(/\s+/g, " ").trim();
  }

  async function loadSearchIndex() {
    if (loaded) return posts;
    const response = await fetch(searchIndexUrl);
    if (!response.ok) {
      throw new Error("Could not load search index");
    }
    posts = await response.json();
    loaded = true;
    return posts;
  }

  function buildSnippet(post, queryTerms) {
    const source = post.content || post.excerpt || "";
    const lowerSource = source.toLowerCase();
    let matchIndex = -1;

    for (const term of queryTerms) {
      const idx = lowerSource.indexOf(term);
      if (idx !== -1 && (matchIndex === -1 || idx < matchIndex)) {
        matchIndex = idx;
      }
    }

    if (matchIndex === -1) {
      return post.excerpt || source.slice(0, 240);
    }

    const start = Math.max(0, matchIndex - 70);
    const end = Math.min(source.length, matchIndex + 170);
    let snippet = source.slice(start, end).trim();
    if (start > 0) snippet = "..." + snippet;
    if (end < source.length) snippet += "...";
    return snippet;
  }

  function renderCards(results, queryTerms) {
    grid.innerHTML = results.map((post) => `
      <article class="post-card">
        <p class="post-meta">${escapeHtml(post.date)}</p>
        <h2><a href="${escapeHtml(post.url)}">${escapeHtml(post.title)}</a></h2>
        <p>${escapeHtml(buildSnippet(post, queryTerms))}</p>
        <a class="read-more" href="${escapeHtml(post.url)}">Read post</a>
      </article>
    `).join("");
  }

  function setStatus(message, hidden = false) {
    status.textContent = message;
    status.hidden = hidden;
  }

  async function runSearch() {
    const query = normalize(input.value);
    if (!query) {
      grid.innerHTML = initialHtml;
      setStatus("", true);
      return;
    }

    try {
      const data = await loadSearchIndex();
      const terms = query.split(" ").filter(Boolean);
      const results = data.filter((post) => {
        const haystack = normalize(`${post.title} ${post.content}`);
        return terms.every((term) => haystack.includes(term));
      });

      renderCards(results, terms);
      const label = results.length === 1 ? "result" : "results";
      setStatus(`${results.length} ${label} for "${input.value.trim()}"`);
    } catch (error) {
      setStatus("Search is temporarily unavailable.");
    }
  }

  input.addEventListener("input", runSearch);
})();
"""
    (ASSETS_DIR / "app.js").write_text(app_js, encoding="utf-8")


def excerpt_from_lines(lines: List[str], max_len: int = 260) -> str:
    cleaned_lines = [line for line in lines if line and not is_excerpt_skippable_line(line)]
    joined = " ".join(cleaned_lines[:12]).strip()
    joined = re.sub(r"\s+", " ", joined)
    if len(joined) <= max_len:
        return joined
    return joined[: max_len - 1].rstrip() + "…"


def build() -> None:
    SITE_DIR.mkdir(parents=True, exist_ok=True)
    POSTS_DIR.mkdir(parents=True, exist_ok=True)
    write_assets()
    IMAGES_DIR.mkdir(parents=True, exist_ok=True)

    character_lines = read_docx_paragraph_lines(CHARACTER_LIST_FILE)
    character_defs = parse_character_list(character_lines)
    variant_lookup, variant_pattern = build_variant_lookup(character_defs)

    all_docx = sorted(SOURCE_DIR.glob("*.docx"))
    episode_files = [
        p
        for p in all_docx
        if p.name.lower().startswith("episode_")
    ]
    episode_files = sorted(episode_files, key=lambda p: parse_episode_number(p.name), reverse=True)

    index_posts = []
    for docx_path in episode_files:
        episode_label = parse_episode_label(docx_path.name)
        title = SPECIAL_TITLES.get(episode_label, f"Episode {episode_label}")
        post_slug = f"episode-{episode_label}" if episode_label != "unknown" else docx_path.stem
        post_url = f"./posts/{post_slug}.html"
        post_image_dir = IMAGES_DIR / post_slug
        if post_image_dir.exists():
            shutil.rmtree(post_image_dir)
        post_image_dir.mkdir(parents=True, exist_ok=True)

        blocks = read_docx_blocks(
            docx_path,
            image_output_dir=post_image_dir,
            image_url_prefix=f"../assets/images/{post_slug}",
        )
        for block in blocks:
            if block.get("type") == "paragraph":
                block["text"] = normalize_episode_text(block["text"], episode_label)
        lines = [b["text"] for b in blocks if b.get("type") == "paragraph"]
        article_blocks = [block_to_html(b, variant_lookup, variant_pattern) for b in blocks]
        article_html = "\n        ".join([b for b in article_blocks if b])

        index_posts.append(
            {
                "title": title,
                "date": dt.datetime.fromtimestamp(docx_path.stat().st_mtime).strftime("%b %d, %Y"),
                "url": post_url,
                "excerpt": excerpt_from_lines(lines),
                "content": " ".join(
                    line for line in lines if line and not is_excerpt_skippable_line(line)
                ),
                "article_html": article_html,
                "canonical_name": docx_path.name,
            }
        )

    for post in index_posts:
        post_slug = Path(post["url"]).stem
        page = render_post_html(
            post["title"],
            post["article_html"],
            post["date"],
            post["canonical_name"],
            index_posts,
            post["url"],
        )
        (POSTS_DIR / f"{post_slug}.html").write_text(page, encoding="utf-8")

    (ASSETS_DIR / "search-index.json").write_text(
        json.dumps(
            [
                {
                    "title": post["title"],
                    "date": post["date"],
                    "url": post["url"],
                    "excerpt": post["excerpt"],
                    "content": post["content"],
                }
                for post in index_posts
            ],
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    (SITE_DIR / "index.html").write_text(render_index_html(index_posts), encoding="utf-8")

    print(f"Built {len(index_posts)} posts into {SITE_DIR}")
    print(f"Loaded {len(character_defs)} character definitions.")


if __name__ == "__main__":
    build()
