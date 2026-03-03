#!/usr/bin/env python3
from __future__ import annotations

import datetime as dt
import html
import re
import shutil
import zipfile
from pathlib import Path
from typing import Dict, List, Tuple
from xml.etree import ElementTree as ET


SOURCE_DIR = Path("/Users/humzaiqbal/Downloads/newsletter")
SITE_DIR = Path("/Users/humzaiqbal/Downloads/trash_tales/site")
POSTS_DIR = SITE_DIR / "posts"
ASSETS_DIR = SITE_DIR / "assets"
IMAGES_DIR = ASSETS_DIR / "images"
CHARACTER_LIST_FILE = SOURCE_DIR / "Character List.docx"

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

        last_match = q_matches[-1]
        desc = line[last_match.end() :].strip(" :-.\t")
        if not desc:
            continue

        for m in q_matches:
            key = canonical_alias(m.group(1).strip())
            if key and key not in aliases:
                aliases[key] = desc
    return aliases


def build_variant_lookup(character_defs: Dict[str, str]) -> Tuple[Dict[str, Tuple[str, str]], re.Pattern]:
    alias_lookup: Dict[str, Tuple[str, str]] = {}
    for alias, desc in character_defs.items():
        norm = normalize_alias_lookup(alias)
        alias_lookup[norm] = (alias, desc)
        if not alias.lower().startswith("the "):
            alias_lookup[normalize_alias_lookup(f"the {alias}")] = (alias, desc)
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
    return f"<p>{inject_character_tooltips(line, variant_lookup, variant_pattern)}</p>"


def parse_episode_number(name: str) -> int:
    m = EPISODE_NUMBER.search(name)
    return int(m.group(1)) if m else -1


def parse_episode_label(name: str) -> str:
    m = EPISODE_LABEL.search(name)
    return m.group(1) if m else "unknown"


def render_post_html(
    title: str,
    article_html: str,
    post_date: dt.datetime,
    canonical_name: str,
) -> str:
    date_fmt = post_date.strftime("%b %d, %Y")
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{html.escape(title)} | Trash Tales</title>
  <meta name="description" content="Weekly Trash Tales newsletter archive by Humza Iqbal." />
  <link rel="stylesheet" href="../assets/styles.css" />
</head>
<body>
  <header class="site-header">
    <div class="site-header-inner">
      <a class="brand" href="../index.html">TRASH TALES</a>
      <nav>
        <a href="../index.html">Archive</a>
      </nav>
    </div>
  </header>

  <main class="container">
    <article class="post">
      <p class="post-meta">{date_fmt}</p>
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
  <script src="../assets/app.js"></script>
</body>
</html>
"""


def render_index_html(posts: List[dict]) -> str:
    cards = []
    for p in posts:
        cards.append(
            f"""
      <article class="post-card">
        <p class="post-meta">{html.escape(p['date'])}</p>
        <h2><a href="{html.escape(p['url'])}">{html.escape(p['title'])}</a></h2>
        <p>{html.escape(p['excerpt'])}</p>
        <a class="read-more" href="{html.escape(p['url'])}">Read post</a>
      </article>
"""
        )
    cards_html = "".join(cards)
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Trash Tales Newsletter</title>
  <meta name="description" content="Weekly Trash Tales newsletter archive by Humza Iqbal." />
  <link rel="stylesheet" href="./assets/styles.css" />
</head>
<body>
  <header class="site-header">
    <div class="site-header-inner">
      <a class="brand" href="./index.html">TRASH TALES</a>
      <nav>
        <a href="./index.html">Archive</a>
      </nav>
    </div>
  </header>

  <main class="container">
    <section class="hero">
      <h1>Weekly Newsletter Archive</h1>
      <p>
        Notes, stories, and characters from each week. Hover or tap a character name to see who they are.
      </p>
    </section>
    <section class="archive-grid">
      {cards_html}
    </section>
  </main>
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
  max-width: 860px;
  margin: 0 auto;
  padding: 14px 20px;
  display: flex;
  align-items: center;
  justify-content: space-between;
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
  max-width: 860px;
  margin: 0 auto;
  padding: 30px 20px 64px;
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
    tooltip.innerHTML = "<strong>" + character + "</strong><br>" + description;
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
        lines = [b["text"] for b in blocks if b.get("type") == "paragraph"]
        article_blocks = [block_to_html(b, variant_lookup, variant_pattern) for b in blocks]
        article_html = "\n        ".join([b for b in article_blocks if b])

        mtime = dt.datetime.fromtimestamp(docx_path.stat().st_mtime)
        page = render_post_html(title, article_html, mtime, docx_path.name)
        (POSTS_DIR / f"{post_slug}.html").write_text(page, encoding="utf-8")

        index_posts.append(
            {
                "title": title,
                "date": mtime.strftime("%b %d, %Y"),
                "url": post_url,
                "excerpt": excerpt_from_lines(lines),
            }
        )

    (SITE_DIR / "index.html").write_text(render_index_html(index_posts), encoding="utf-8")

    print(f"Built {len(index_posts)} posts into {SITE_DIR}")
    print(f"Loaded {len(character_defs)} character definitions.")


if __name__ == "__main__":
    build()
