(() => {
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
