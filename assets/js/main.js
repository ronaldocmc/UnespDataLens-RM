(function () {
  "use strict";

  const normalize = (value) =>
    (value || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase();

  const progress = document.querySelector(".reading-progress span");
  if (progress) {
    const updateProgress = () => {
      const root = document.documentElement;
      const distance = root.scrollHeight - root.clientHeight;
      const ratio = distance > 0 ? (root.scrollTop / distance) * 100 : 0;
      progress.style.width = `${Math.min(100, Math.max(0, ratio))}%`;
    };
    document.addEventListener("scroll", updateProgress, { passive: true });
    updateProgress();
  }

  const cardFilter = document.querySelector("[data-card-filter]");
  if (cardFilter) {
    const cards = Array.from(document.querySelectorAll(".lexicon-grid .card"));
    cardFilter.addEventListener("input", () => {
      const query = normalize(cardFilter.value.trim());
      cards.forEach((card) => {
        card.hidden = query && !normalize(card.textContent).includes(query);
      });
    });
  }

  const searchInput = document.getElementById("search-query");
  const searchButton = document.getElementById("search-button");
  const searchResults = document.getElementById("search-results");
  const searchMeta = document.getElementById("search-meta");

  if (searchInput && searchResults && window.UNESP_DATALENS_INDEX) {
    const runSearch = () => {
      const query = normalize(searchInput.value.trim());
      const tokens = query.split(/\s+/).filter(Boolean);
      const ranked = window.UNESP_DATALENS_INDEX
        .map((record) => {
          const title = normalize(record.title);
          const haystack = normalize(`${record.title} ${record.type} ${record.text}`);
          if (!tokens.every((token) => haystack.includes(token))) return null;
          let score = 0;
          if (title === query) score += 100;
          if (title.startsWith(query)) score += 50;
          tokens.forEach((token) => {
            if (title.includes(token)) score += 12;
          });
          return { record, score };
        })
        .filter(Boolean)
        .sort((a, b) => b.score - a.score || a.record.title.localeCompare(b.record.title, "pt-BR"));

      searchMeta.textContent = query
        ? `${ranked.length} resultado${ranked.length === 1 ? "" : "s"}`
        : "Digite dois ou mais caracteres para pesquisar.";
      if (query.length < 2) {
        searchResults.innerHTML = "";
        return;
      }
      searchResults.innerHTML =
        ranked
          .slice(0, 120)
          .map(
            ({ record }) =>
              `<a class="search-result" href="${record.url}"><strong>${escapeHtml(
                record.title
              )}</strong><span class="result-type">${escapeHtml(
                record.type
              )}</span><p>${escapeHtml(record.text.slice(0, 300))}</p></a>`
          )
          .join("") || '<div class="box">Nenhum resultado encontrado. Tente um termo mais amplo.</div>';
    };

    searchInput.addEventListener("input", runSearch);
    searchInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        runSearch();
      }
    });
    if (searchButton) searchButton.addEventListener("click", runSearch);
    const params = new URLSearchParams(location.search);
    if (params.get("q")) {
      searchInput.value = params.get("q");
      runSearch();
    }
  }

  function escapeHtml(value) {
    return String(value)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }
})();
