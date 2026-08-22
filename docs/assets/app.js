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

  function buildSnippet(post, query) {
    const source = post.content || post.excerpt || "";
    const matchIndex = normalize(source).indexOf(query);

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

  function renderCards(results, query) {
    grid.innerHTML = results.map((post) => `
      <article class="post-card">
        <p class="post-meta">${escapeHtml(post.date)}</p>
        <h2><a href="${escapeHtml(post.url)}">${escapeHtml(post.title)}</a></h2>
        <p>${escapeHtml(buildSnippet(post, query))}</p>
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
      const results = data.filter((post) => {
        const haystack = normalize(`${post.title} ${post.content}`);
        return haystack.includes(query);
      });

      renderCards(results, query);
      const label = results.length === 1 ? "result" : "results";
      setStatus(`${results.length} ${label} for "${input.value.trim()}"`);
    } catch (error) {
      setStatus("Search is temporarily unavailable.");
    }
  }

  input.addEventListener("input", runSearch);
})();

(() => {
  const allCards = window.GOVERNMENT_FLASHCARDS || [];
  const cardButton = document.getElementById("flashcard");
  const cardText = document.getElementById("flashcard-text");
  const cardSide = document.getElementById("flashcard-side");
  const cardDeck = document.getElementById("flashcard-deck");
  const deckFilter = document.getElementById("deck-filter");
  const progress = document.getElementById("quiz-progress");
  const list = document.getElementById("quiz-card-list");
  const shuffleButton = document.getElementById("shuffle-cards");
  const reviewButtons = [
    document.getElementById("again-card"),
    document.getElementById("good-card"),
    document.getElementById("easy-card"),
  ];

  if (!allCards.length || !cardButton || !cardText || !cardSide || !cardDeck || !deckFilter || !progress || !list) return;

  let cards = [...allCards];
  let index = 0;
  let showingBack = false;

  function escapeHtml(value) {
    return value
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  function filteredCards() {
    const selectedDeck = deckFilter.value;
    if (selectedDeck === "all") return [...allCards];
    return allCards.filter((card) => card.deck === selectedDeck);
  }

  function renderList() {
    list.innerHTML = cards.map((card) => `
      <article class="quiz-card-summary">
        <p class="post-meta">${escapeHtml(card.deck)}</p>
        <h3>${escapeHtml(card.front)}</h3>
        <p>${escapeHtml(card.back)}</p>
      </article>
    `).join("");
  }

  function renderCard() {
    if (!cards.length) {
      cardDeck.textContent = "";
      cardSide.textContent = "No cards";
      cardText.textContent = "No cards match this deck.";
      progress.textContent = "";
      list.innerHTML = "";
      return;
    }

    const card = cards[index];
    cardDeck.textContent = card.deck;
    cardSide.textContent = showingBack ? "Answer" : "Question";
    cardText.textContent = showingBack ? card.back : card.front;
    progress.textContent = `Card ${index + 1} of ${cards.length}`;
    renderList();
  }

  function nextCard() {
    if (!cards.length) return;
    index = (index + 1) % cards.length;
    showingBack = false;
    renderCard();
  }

  function shuffleCards() {
    for (let i = cards.length - 1; i > 0; i -= 1) {
      const j = Math.floor(Math.random() * (i + 1));
      [cards[i], cards[j]] = [cards[j], cards[i]];
    }
    index = 0;
    showingBack = false;
    renderCard();
  }

  cardButton.addEventListener("click", () => {
    showingBack = !showingBack;
    renderCard();
  });

  deckFilter.addEventListener("change", () => {
    cards = filteredCards();
    index = 0;
    showingBack = false;
    renderCard();
  });

  shuffleButton?.addEventListener("click", shuffleCards);
  reviewButtons.forEach((button) => button?.addEventListener("click", nextCard));

  renderCard();
})();
