const SOURCE_FILES = [
  { id: "bbook", url: "bbook.json", label: "Big Book" },
  { id: "twlvxtwlv", url: "twlvxtwlv.json", label: "Twelve Steps and Twelve Traditions" },
  { id: "daily", url: "daily.json", label: "Daily Reflections" }
];

const SOURCE_ORDER = ["steps", "bbook", "traditions", "daily"];

const SOURCE_HINTS = {
  steps: "Upload twlvxtwlv.json to see the Steps.",
  bbook: "Upload bbook.json to see the Big Book.",
  traditions: "Upload twlvxtwlv.json to see the Traditions.",
  daily: "Upload daily.json to see Daily Reflections."
};

const MONTHS = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December"
];

const PRINCIPLE_TERMS = [
  "honesty",
  "hope",
  "faith",
  "courage",
  "integrity",
  "willingness",
  "humility",
  "love",
  "service",
  "tolerance",
  "patience",
  "perseverance",
  "forgiveness",
  "compassion",
  "kindness",
  "gratitude",
  "responsibility",
  "acceptance",
  "open-mindedness",
  "open mindedness"
];

const HIGHLIGHT_RULES = [
  {
    id: "principle",
    className: "principle",
    terms: PRINCIPLE_TERMS,
    wordBoundary: true,
    priority: 1
  }
];

const CACHE_DB_NAME = "aa-study-library";
const CACHE_STORE = "sources";

const state = {
  sources: [],
  sourceById: new Map(),
  activeSourceId: null,
  activeSectionBySource: new Map(),
  activeMonth: null,
  searchTerm: "",
  paragraphIndex: [],
  openNavId: null,
  dailyQuotes: []
};

const elements = {};

document.addEventListener("DOMContentLoaded", () => {
  cacheElements();
  bindEvents();
  loadSources();
});

function cacheElements() {
  elements.statusLine = document.getElementById("statusLine");
  elements.searchForm = document.getElementById("searchForm");
  elements.searchInput = document.getElementById("searchInput");
  elements.clearSearchBtn = document.getElementById("clearSearchBtn");
  elements.searchSummary = document.getElementById("searchSummary");
  elements.resultsList = document.getElementById("resultsList");
  elements.refreshBtn = document.getElementById("refreshBtn");
  elements.folderInput = document.getElementById("folderInput");
  elements.fileInput = document.getElementById("fileInput");
  elements.localFallback = document.getElementById("localFallback");
  elements.localFallbackMessage = document.getElementById("localFallbackMessage");
  elements.activeKicker = document.getElementById("activeKicker");
  elements.activeTitle = document.getElementById("activeTitle");
  elements.activeMeta = document.getElementById("activeMeta");
  elements.content = document.getElementById("content");
  elements.nav = document.getElementById("topNav");
  elements.dropdownArea = document.getElementById("dropdownArea");
  elements.navButtons = {
    steps: document.getElementById("navSteps"),
    bbook: document.getElementById("navBigBook"),
    traditions: document.getElementById("navTraditions"),
    daily: document.getElementById("navDaily")
  };
  elements.dropdownPanels = {
    steps: document.getElementById("dropdownSteps"),
    bbook: document.getElementById("dropdownBigBook"),
    traditions: document.getElementById("dropdownTraditions"),
    daily: document.getElementById("dropdownDaily")
  };
  elements.dropdownLists = {
    steps: document.getElementById("stepsList"),
    bbook: document.getElementById("bbookList"),
    traditions: document.getElementById("traditionsList")
  };
  elements.dailyMonthStrip = document.getElementById("dailyMonthStrip");
  elements.dailyDayList = document.getElementById("dailyDayList");
}

function bindEvents() {
  if (elements.searchForm) {
    elements.searchForm.addEventListener("submit", onSearchSubmit);
  }
  if (elements.clearSearchBtn) {
    elements.clearSearchBtn.addEventListener("click", clearSearch);
  }
  if (elements.refreshBtn) {
    elements.refreshBtn.addEventListener("click", () => loadSources(true));
  }
  if (elements.folderInput) {
    elements.folderInput.addEventListener("change", handleFileSelection);
  }
  if (elements.fileInput) {
    elements.fileInput.addEventListener("change", handleFileSelection);
  }

  Object.entries(elements.navButtons).forEach(([id, button]) => {
    if (!button) return;
    button.addEventListener("click", (event) => {
      event.stopPropagation();
      onNavClick(id);
    });
  });

  document.addEventListener("click", (event) => {
    if (!state.openNavId) return;
    const target = event.target;
    if (elements.nav && elements.nav.contains(target)) return;
    if (elements.dropdownArea && elements.dropdownArea.contains(target)) return;
    closeDropdown();
  });

  window.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      closeDropdown();
    }
  });
}

function onNavClick(sourceId) {
  const source = state.sourceById.get(sourceId);
  if (!source) return;
  setActiveSource(sourceId);
  toggleDropdown(sourceId);
}

function toggleDropdown(sourceId) {
  if (state.openNavId === sourceId) {
    closeDropdown();
    return;
  }
  state.openNavId = sourceId;
  updateDropdownVisibility();
}

function closeDropdown() {
  state.openNavId = null;
  updateDropdownVisibility();
}

function updateDropdownVisibility() {
  Object.entries(elements.dropdownPanels).forEach(([id, panel]) => {
    if (!panel) return;
    const isOpen = state.openNavId === id;
    panel.classList.toggle("open", isOpen);
    const button = elements.navButtons[id];
    if (button) {
      button.classList.toggle("open", isOpen);
      button.setAttribute("aria-expanded", isOpen ? "true" : "false");
      if (panel.id) {
        button.setAttribute("aria-controls", panel.id);
      }
    }
  });
}

async function loadSources(forceReload) {
  setStatus("Loading library...");
  hideLocalFallback();

  const [cachedEntries, results] = await Promise.all([
    getCachedSources(),
    Promise.allSettled(SOURCE_FILES.map((source) => loadSourceFile(source, forceReload)))
  ]);

  const cachedSources = [];
  cachedEntries.forEach((entry) => {
    const normalized = normalizeSource(entry.id, entry.data, entry.label);
    pushSources(cachedSources, normalized);
  });

  const fetchedSources = [];
  const cacheWrites = [];
  results.forEach((result) => {
    if (result.status !== "fulfilled") return;
    const payload = result.value;
    pushSources(fetchedSources, payload.normalized);
    cacheWrites.push(cacheSource(payload.id, payload.label, payload.data));
  });
  if (cacheWrites.length) {
    await Promise.allSettled(cacheWrites);
  }

  const sources = mergeSources(cachedSources, fetchedSources);

  if (!sources.length) {
    if (location.protocol === "file:") {
      showLocalFallback(
        "This page is opened from disk, so the browser blocks loading JSON files directly. Choose the reader folder or select the JSON files."
      );
    }
    clearLibrary();
    return;
  }

  applySources(sources);
}

function clearLibrary() {
  state.sources = [];
  state.sourceById = new Map();
  state.activeSourceId = null;
  state.activeSectionBySource = new Map();
  state.activeMonth = null;
  state.paragraphIndex = [];
  state.dailyQuotes = [];
  elements.content.innerHTML = "";
  elements.resultsList.innerHTML = "";
  renderSearchSummary(0, 0);
  renderEmptyReader();
  renderNav();
  renderDropdowns();
  renderStatus();
}

async function loadSourceFile(source, forceReload) {
  const url = forceReload ? `${source.url}?ts=${Date.now()}` : source.url;
  const response = await fetch(url, { cache: "no-store" });
  if (!response.ok) throw new Error("File load failed");
  const data = await response.json();
  return {
    id: source.id,
    label: source.label,
    data,
    normalized: normalizeSource(source.id, data, source.label)
  };
}

function normalizeSource(id, data, fallbackTitle) {
  if (Array.isArray(data)) {
    return normalizeDailySource(id, data, fallbackTitle);
  }

  if (data && (data.steps || data.traditions || data.foreword)) {
    return normalizeTwelveSource(id, data, fallbackTitle);
  }

  if (data && Array.isArray(data.sections)) {
    return normalizeSectionSource(id, data, fallbackTitle);
  }

  return {
    id,
    title: fallbackTitle || id,
    kind: "unknown",
    metadata: "",
    author: "",
    sections: []
  };
}

function normalizeSectionSource(id, data, fallbackTitle) {
  const sections = (data.sections || []).map((section, index) => {
    const heading = section.title || section.heading || `Section ${index + 1}`;
    const content = Array.isArray(section.content)
      ? section.content
      : Array.isArray(section.paragraphs)
      ? section.paragraphs
      : [];

    return {
      id: section.id || `${id}-section-${index + 1}`,
      heading,
      paragraphs: content.filter(Boolean),
      meta: {
        type: section.type || "section",
        number: section.number || null,
        subtitle: section.subtitle || ""
      }
    };
  });

  const source = {
    id,
    title: data.title || fallbackTitle || id,
    kind: "book",
    metadata: data.metadata || "",
    author: data.author || "",
    sections
  };

  attachSectionIndex(source);
  return source;
}

function normalizeTwelveSource(id, data, fallbackTitle) {
  const steps = [];
  const traditions = [];

  if (data.foreword) {
    steps.push(sectionFromEntry(data.foreword, "foreword"));
  }

  if (Array.isArray(data.steps)) {
    data.steps.forEach((step) => {
      steps.push(sectionFromEntry(step, "step"));
    });
  }

  if (Array.isArray(data.traditions)) {
    data.traditions.forEach((tradition) => {
      traditions.push(sectionFromEntry(tradition, "tradition"));
    });
  }

  const sources = [];
  if (steps.length) {
    sources.push(makeTwelveSource("steps", "Twelve Steps", data, steps));
  }
  if (traditions.length) {
    sources.push(makeTwelveSource("traditions", "Twelve Traditions", data, traditions));
  }

  if (sources.length) return sources;

  return {
    id,
    title: fallbackTitle || id,
    kind: "book",
    metadata: data.metadata || "",
    author: data.author || "",
    sections: []
  };
}

function makeTwelveSource(id, title, data, sections) {
  const source = {
    id,
    title,
    kind: "book",
    metadata: data.metadata || "",
    author: data.author || "",
    sections
  };
  attachSectionIndex(source);
  return source;
}

function sectionFromEntry(entry, type) {
  const heading = entry.title || entry.name || "Untitled";
  const paragraphs = [];
  if (entry.subtitle) {
    paragraphs.push(entry.subtitle);
  }
  if (Array.isArray(entry.content)) {
    paragraphs.push(...entry.content);
  }

  return {
    id: entry.id || slugify(heading),
    heading,
    paragraphs: paragraphs.filter(Boolean),
    meta: {
      type,
      number: entry.number || null,
      subtitle: entry.subtitle || ""
    }
  };
}

function normalizeDailySource(id, data, fallbackTitle) {
  const sections = data.map((entry, index) => {
    const heading = `${entry.date || `Day ${index + 1}`} - ${entry.title || ""}`.trim();
    const paragraphs = [entry.quote, entry.reflection].filter(Boolean);
    return {
      id: slugify(`${entry.month || "month"}-${entry.day || index + 1}`),
      heading,
      paragraphs,
      meta: {
        type: "daily",
        date: entry.date || "",
        title: entry.title || "",
        quote: entry.quote || "",
        source: entry.source || "",
        reflection: entry.reflection || "",
        month: entry.month || "",
        day: entry.day || null,
        pageIndex: entry.page_index || null
      }
    };
  });

  const source = {
    id,
    title: fallbackTitle || "Daily Reflections",
    kind: "daily",
    metadata: "",
    author: "",
    sections
  };

  attachSectionIndex(source);
  return source;
}

function attachSectionIndex(source) {
  source.sectionById = new Map();
  source.sections.forEach((section, index) => {
    section.index = index;
    source.sectionById.set(section.id, section);
  });
}

function applySources(sources) {
  state.sources = sources;
  state.sourceById = new Map(sources.map((source) => [source.id, source]));
  buildDailyQuoteIndex();

  if (!state.activeSourceId || !state.sourceById.has(state.activeSourceId)) {
    state.activeSourceId = getFirstAvailableSourceId();
  }

  state.activeSectionBySource.forEach((_value, key) => {
    if (!state.sourceById.has(key)) {
      state.activeSectionBySource.delete(key);
    }
  });

  state.sources.forEach((source) => {
    ensureActiveSection(source);
  });

  buildParagraphIndex();
  renderActiveSource();

  if (state.searchTerm) {
    runSearch(state.searchTerm);
  } else {
    renderSearchSummary(0, 0);
    elements.resultsList.innerHTML = "";
  }

  renderStatus();
}

function pushSources(target, value) {
  if (!value) return;
  if (Array.isArray(value)) {
    value.forEach((item) => target.push(item));
    return;
  }
  target.push(value);
}

function mergeSources(cachedSources, fetchedSources) {
  const merged = new Map();
  cachedSources.forEach((source) => {
    if (!source || !source.id) return;
    merged.set(source.id, source);
  });
  fetchedSources.forEach((source) => {
    if (!source || !source.id) return;
    merged.set(source.id, source);
  });

  const ordered = [];
  SOURCE_ORDER.forEach((id) => {
    if (merged.has(id)) {
      ordered.push(merged.get(id));
      merged.delete(id);
    }
  });
  merged.forEach((source) => ordered.push(source));
  return ordered;
}

function getFirstAvailableSourceId() {
  for (const id of SOURCE_ORDER) {
    if (state.sourceById.has(id)) return id;
  }
  return state.sources[0] ? state.sources[0].id : null;
}

function ensureActiveSection(source) {
  const stored = state.activeSectionBySource.get(source.id);
  if (stored && source.sectionById.has(stored)) return stored;
  const fallback =
    source.kind === "daily" ? getDefaultDailyId(source) : source.sections[0]?.id || null;
  if (fallback) {
    state.activeSectionBySource.set(source.id, fallback);
  }
  return fallback;
}

function buildParagraphIndex() {
  state.paragraphIndex = [];
  state.sources.forEach((source, sourceIndex) => {
    const sourceKey = `s${sourceIndex}`;
    source.key = sourceKey;
    source.sections.forEach((section) => {
      section.paragraphs.forEach((text, paragraphIndex) => {
        state.paragraphIndex.push({
          sourceId: source.id,
          sourceTitle: source.title,
          sectionId: section.id,
          heading: section.heading,
          text,
          sectionIndex: section.index,
          paragraphIndex,
          domId: buildParagraphId(sourceKey, section.index, paragraphIndex)
        });
      });
    });
  });
}

function buildParagraphId(sourceKey, sectionIndex, paragraphIndex) {
  return `p-${sourceKey}-${sectionIndex}-${paragraphIndex}`;
}

function setActiveSource(sourceId) {
  if (!sourceId || !state.sourceById.has(sourceId)) return;
  state.activeSourceId = sourceId;
  renderActiveSource();
}

function renderActiveSource() {
  const source = state.sourceById.get(state.activeSourceId);
  if (!source) {
    renderEmptyReader();
    renderNav();
    renderDropdowns();
    return;
  }

  const sectionId = ensureActiveSection(source);
  const section = sectionId ? source.sectionById.get(sectionId) : null;

  if (source.kind === "daily") {
    renderDailySection(source, section);
  } else {
    renderBookSection(source, section);
  }

  renderNav();
  renderDropdowns();
}

function renderEmptyReader() {
  elements.activeKicker.textContent = "";
  elements.activeTitle.textContent = "Upload the library files";
  elements.activeMeta.textContent = "Steps, Big Book, Traditions, and Daily Reflections.";
  elements.content.innerHTML = "";
}

function renderBookSection(source, section) {
  elements.activeKicker.textContent = source.title;
  elements.activeTitle.textContent = section?.heading || source.title;
  elements.activeMeta.textContent = source.author || "";
  elements.content.innerHTML = "";

  if (!section) {
    const empty = document.createElement("div");
    empty.className = "panel-note";
    empty.textContent = "No content available.";
    elements.content.appendChild(empty);
    return;
  }

  const wrapper = document.createElement("section");
  wrapper.className = "section-block";

  const subtitle = section.meta?.subtitle;
  const paragraphs = [...section.paragraphs];
  let paragraphOffset = 0;
  if (subtitle) {
    const sub = document.createElement("p");
    sub.className = "section-subtitle";
    if (paragraphs[0] === subtitle) {
      sub.id = buildParagraphId(source.key, section.index, 0);
      paragraphs.shift();
      paragraphOffset = 1;
    }
    appendHighlightedText(sub, subtitle, state.searchTerm, { includeCrossrefs: true });
    wrapper.appendChild(sub);
  }

  paragraphs.forEach((text, paragraphIndex) => {
    const paragraph = document.createElement("p");
    paragraph.className = "para";
    paragraph.id = buildParagraphId(
      source.key,
      section.index,
      paragraphIndex + paragraphOffset
    );
    appendHighlightedText(paragraph, text, state.searchTerm, { includeCrossrefs: true });
    wrapper.appendChild(paragraph);
  });

  elements.content.appendChild(wrapper);
}

function renderDailySection(source, section) {
  if (!section) {
    renderBookSection(source, section);
    return;
  }

  const meta = section.meta || {};
  if (meta.month) {
    state.activeMonth = meta.month;
  }

  elements.activeKicker.textContent = meta.date || "Daily Reflection";
  elements.activeTitle.textContent = meta.title || source.title;
  elements.activeMeta.textContent = meta.source || "";

  elements.content.innerHTML = "";

  const wrapper = document.createElement("section");
  wrapper.className = "section-block";

  let paragraphIndex = 0;
  if (meta.quote) {
    const quote = document.createElement("blockquote");
    quote.id = buildParagraphId(source.key, section.index, paragraphIndex);
    appendHighlightedText(quote, meta.quote, state.searchTerm, { includeCrossrefs: false });
    wrapper.appendChild(quote);
    paragraphIndex += 1;
  }

  if (meta.source) {
    const sourceLine = document.createElement("div");
    sourceLine.className = "quote-source";
    sourceLine.textContent = meta.source;
    wrapper.appendChild(sourceLine);
  }

  if (meta.reflection) {
    const paragraph = document.createElement("p");
    paragraph.id = buildParagraphId(source.key, section.index, paragraphIndex);
    appendHighlightedText(paragraph, meta.reflection, state.searchTerm, {
      includeCrossrefs: false
    });
    wrapper.appendChild(paragraph);
  }

  elements.content.appendChild(wrapper);
}

function getDefaultDailyId(source) {
  const now = new Date();
  const month = MONTHS[now.getMonth()];
  const day = now.getDate();
  const match = source.sections.find(
    (section) => section.meta?.month === month && section.meta?.day === day
  );
  return match ? match.id : source.sections[0]?.id || null;
}

function renderNav() {
  SOURCE_ORDER.forEach((id) => {
    const button = elements.navButtons[id];
    if (!button) return;
    const source = state.sourceById.get(id);
    const isActive = state.activeSourceId === id;
    button.disabled = !source;
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-disabled", button.disabled ? "true" : "false");
    const hint = SOURCE_HINTS[id];
    button.title = !source && hint ? hint : "";
  });

  if (state.openNavId && !state.sourceById.has(state.openNavId)) {
    state.openNavId = null;
  }
  updateDropdownVisibility();
}

function renderDropdowns() {
  renderTocDropdown("steps");
  renderTocDropdown("bbook");
  renderTocDropdown("traditions");
  renderDailyDropdown();
}

function renderTocDropdown(sourceId) {
  const list = elements.dropdownLists[sourceId];
  if (!list) return;
  list.innerHTML = "";

  const source = state.sourceById.get(sourceId);
  if (!source || !source.sections.length) {
    const empty = document.createElement("div");
    empty.className = "dropdown-empty";
    empty.textContent = SOURCE_HINTS[sourceId] || "Upload a source to see sections.";
    list.appendChild(empty);
    return;
  }

  const activeSectionId = ensureActiveSection(source);

  source.sections.forEach((section, index) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "dropdown-item";
    button.textContent = formatSectionLabel(section);
    button.style.setProperty("--i", index);
    if (section.id === activeSectionId) {
      button.classList.add("active");
    }
    button.addEventListener("click", () => {
      state.activeSectionBySource.set(source.id, section.id);
      setActiveSource(source.id);
      closeDropdown();
    });
    list.appendChild(button);
  });
}

function renderDailyDropdown() {
  const source = state.sourceById.get("daily");
  if (!elements.dailyMonthStrip || !elements.dailyDayList) return;
  elements.dailyMonthStrip.innerHTML = "";
  elements.dailyDayList.innerHTML = "";

  if (!source || !source.sections.length) {
    const empty = document.createElement("div");
    empty.className = "dropdown-empty";
    empty.textContent = SOURCE_HINTS.daily;
    elements.dailyDayList.appendChild(empty);
    return;
  }

  const activeSectionId = ensureActiveSection(source);
  const activeSection = activeSectionId ? source.sectionById.get(activeSectionId) : null;
  const activeMonth = activeSection?.meta?.month || state.activeMonth || MONTHS[0];
  state.activeMonth = activeMonth;

  MONTHS.forEach((month) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "month-btn";
    button.textContent = month.slice(0, 3);
    if (month === activeMonth) {
      button.classList.add("active");
    }
    button.addEventListener("click", () => {
      state.activeMonth = month;
      const monthEntries = source.sections
        .filter((section) => section.meta?.month === month)
        .sort((a, b) => (a.meta.day || 0) - (b.meta.day || 0));
      if (monthEntries[0]) {
        state.activeSectionBySource.set(source.id, monthEntries[0].id);
        setActiveSource(source.id);
      }
    });
    elements.dailyMonthStrip.appendChild(button);
  });

  const entries = source.sections
    .filter((section) => section.meta?.month === activeMonth)
    .sort((a, b) => (a.meta.day || 0) - (b.meta.day || 0));

  entries.forEach((section) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "day-item";
    if (section.id === activeSectionId) {
      button.classList.add("active");
    }

    const dayLabel = document.createElement("strong");
    dayLabel.textContent = section.meta?.date || section.heading;
    const title = document.createElement("span");
    title.textContent = section.meta?.title || section.heading;

    button.appendChild(dayLabel);
    button.appendChild(title);
    button.addEventListener("click", () => {
      state.activeSectionBySource.set(source.id, section.id);
      setActiveSource(source.id);
      closeDropdown();
    });
    elements.dailyDayList.appendChild(button);
  });
}

function formatSectionLabel(section) {
  if (!section) return "Untitled";
  return section.heading || "Untitled";
}

function appendHighlightedText(node, text, term, options = {}) {
  const safeText = text ? String(text) : "";
  node.textContent = "";
  const matches = buildHighlightMatches(safeText, term, options);
  if (!matches.length) {
    node.textContent = safeText;
    return;
  }

  let cursor = 0;
  matches.forEach((match) => {
    if (match.start > cursor) {
      node.appendChild(document.createTextNode(safeText.slice(cursor, match.start)));
    }
    const mark = document.createElement("mark");
    mark.className = match.className;
    if (match.title) {
      mark.title = match.title;
    }
    mark.textContent = safeText.slice(match.start, match.end);
    node.appendChild(mark);
    cursor = match.end;
  });

  if (cursor < safeText.length) {
    node.appendChild(document.createTextNode(safeText.slice(cursor)));
  }
}

function onSearchSubmit(event) {
  event.preventDefault();
  const term = elements.searchInput.value.trim();
  if (!term) {
    clearSearch();
    return;
  }
  runSearch(term);
}

function runSearch(term) {
  state.searchTerm = term;
  const result = searchParagraphs(term);
  renderSearchSummary(result.totalHits, result.results.length);
  renderResults(result.results, term);
  renderActiveSource();
}

function clearSearch() {
  state.searchTerm = "";
  elements.searchInput.value = "";
  elements.resultsList.innerHTML = "";
  renderSearchSummary(0, 0);
  renderActiveSource();
}

function searchParagraphs(term) {
  const needle = term.toLowerCase();
  const results = [];
  let totalHits = 0;

  state.paragraphIndex.forEach((item) => {
    const count = countOccurrences(item.text, needle);
    if (count > 0) {
      totalHits += count;
      results.push({
        ...item,
        count,
        snippet: makeSnippet(item.text, term)
      });
    }
  });

  return { results, totalHits };
}

function countOccurrences(text, needleLower) {
  if (!needleLower) return 0;
  const lower = String(text || "").toLowerCase();
  let count = 0;
  let idx = 0;
  while ((idx = lower.indexOf(needleLower, idx)) !== -1) {
    count += 1;
    idx += needleLower.length;
  }
  return count;
}

function makeSnippet(text, term) {
  const safeText = String(text || "");
  const lower = safeText.toLowerCase();
  const needle = term.toLowerCase();
  const idx = lower.indexOf(needle);
  if (idx === -1) {
    return safeText.slice(0, 140);
  }
  const context = 60;
  const start = Math.max(0, idx - context);
  const end = Math.min(safeText.length, idx + needle.length + context);
  let snippet = safeText.slice(start, end).trim();
  if (start > 0) snippet = "..." + snippet;
  if (end < safeText.length) snippet = snippet + "...";
  return snippet;
}

function renderSearchSummary(totalHits, sections) {
  if (!state.searchTerm) {
    elements.searchSummary.textContent = "Search across all sources.";
    return;
  }
  const hitLabel = totalHits === 1 ? "hit" : "hits";
  const sectionLabel = sections === 1 ? "paragraph" : "paragraphs";
  elements.searchSummary.textContent = `${totalHits} ${hitLabel} in ${sections} ${sectionLabel}.`;
}

function renderResults(results, term) {
  const list = elements.resultsList;
  list.innerHTML = "";
  if (!term) return;

  if (!results.length) {
    const empty = document.createElement("div");
    empty.className = "panel-note";
    empty.textContent = "No matches found.";
    list.appendChild(empty);
    return;
  }

  const maxResults = 150;
  results.slice(0, maxResults).forEach((result, index) => {
    const item = document.createElement("button");
    item.type = "button";
    item.className = "result-item";
    item.style.setProperty("--i", index);

    const title = document.createElement("div");
    title.className = "result-title";
    const heading = result.heading ? ` - ${result.heading}` : "";
    title.textContent = `${result.sourceTitle}${heading} (${result.count})`;

    const snippet = document.createElement("div");
    snippet.className = "result-snippet";
    appendHighlightedText(snippet, result.snippet, term, { includeCrossrefs: false });

    item.appendChild(title);
    item.appendChild(snippet);
    item.addEventListener("click", () => navigateToResult(result));

    list.appendChild(item);
  });

  if (results.length > maxResults) {
    const note = document.createElement("div");
    note.className = "panel-note";
    note.textContent = `Showing first ${maxResults} matches.`;
    list.appendChild(note);
  }
}

function buildHighlightMatches(text, term, options = {}) {
  const candidates = [];
  const used = [];

  if (term) {
    const searchRegex = buildRegex([term], false);
    if (searchRegex) {
      collectMatches(text, searchRegex, "hit", 3, candidates);
    }
  }

  if (options.includeCrossrefs && state.dailyQuotes.length) {
    collectCrossRefMatches(text, state.dailyQuotes, candidates);
  }

  HIGHLIGHT_RULES.forEach((rule) => {
    const regex = buildRegex(rule.terms, rule.wordBoundary);
    if (!regex) return;
    collectMatches(text, regex, rule.className, rule.priority, candidates);
  });

  if (!candidates.length) return [];
  candidates.sort((a, b) => {
    if (b.priority !== a.priority) return b.priority - a.priority;
    if (a.start !== b.start) return a.start - b.start;
    return b.end - b.start - (a.end - a.start);
  });

  const final = [];
  candidates.forEach((match) => {
    if (hasOverlap(match, used)) return;
    final.push(match);
    insertRange(used, match);
  });

  final.sort((a, b) => a.start - b.start);
  return final;
}

function collectMatches(text, regex, className, priority, output) {
  let match;
  while ((match = regex.exec(text)) !== null) {
    if (!match[0]) continue;
    output.push({
      start: match.index,
      end: match.index + match[0].length,
      className,
      priority
    });
    if (match.index === regex.lastIndex) {
      regex.lastIndex += 1;
    }
  }
}

function collectCrossRefMatches(text, quotes, output) {
  const haystack = String(text || "");
  const haystackLower = haystack.toLowerCase();
  quotes.forEach((entry) => {
    const quote = entry.text;
    if (!quote) return;
    const needleLower = buildCaseInsensitiveNeedle(quote);
    if (!needleLower) return;
    let idx = 0;
    while ((idx = haystackLower.indexOf(needleLower, idx)) !== -1) {
      output.push({
        start: idx,
        end: idx + needleLower.length,
        className: "crossref",
        title: entry.date ? `Daily Reflection: ${entry.date}` : "Daily Reflection",
        priority: 2
      });
      idx += needleLower.length;
    }
  });
}

function buildCaseInsensitiveNeedle(text) {
  const trimmed = String(text || "").trim();
  if (!trimmed) return "";
  return trimmed.toLowerCase();
}

function buildRegex(terms, wordBoundary) {
  const filtered = (terms || []).map((term) => String(term || "").trim()).filter(Boolean);
  if (!filtered.length) return null;
  const pattern = filtered.map(escapeRegex).join("|");
  const boundary = wordBoundary ? "\\b" : "";
  return new RegExp(`${boundary}(?:${pattern})${boundary}`, "gi");
}

function escapeRegex(text) {
  return String(text || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function hasOverlap(match, ranges) {
  return ranges.some((range) => match.start < range.end && match.end > range.start);
}

function insertRange(ranges, match) {
  ranges.push({ start: match.start, end: match.end });
  ranges.sort((a, b) => a.start - b.start);
  for (let i = 0; i < ranges.length - 1; i += 1) {
    const current = ranges[i];
    const next = ranges[i + 1];
    if (current.end >= next.start) {
      current.end = Math.max(current.end, next.end);
      ranges.splice(i + 1, 1);
      i -= 1;
    }
  }
}

function navigateToResult(result) {
  const source = state.sourceById.get(result.sourceId);
  if (!source) return;
  state.activeSectionBySource.set(source.id, result.sectionId);
  if (source.kind === "daily") {
    const section = source.sectionById.get(result.sectionId);
    if (section?.meta?.month) {
      state.activeMonth = section.meta.month;
    }
  }
  setActiveSource(source.id);
  closeDropdown();

  setTimeout(() => {
    scrollToParagraph(result.domId);
  }, 0);
}

function scrollToParagraph(domId) {
  const el = document.getElementById(domId);
  if (!el) return;
  el.classList.add("spotlight");
  el.scrollIntoView({ behavior: "smooth", block: "start" });
  window.setTimeout(() => {
    el.classList.remove("spotlight");
  }, 1400);
}

function setStatus(message) {
  if (elements.statusLine) {
    elements.statusLine.textContent = message;
  }
}

function renderStatus() {
  const sourceCount = state.sources.length;
  const sourceLabel = sourceCount === 1 ? "source" : "sources";
  setStatus(sourceCount ? `${sourceCount} ${sourceLabel} loaded.` : "No sources loaded.");
}

function showLocalFallback(message) {
  if (!elements.localFallback) return;
  elements.localFallback.classList.remove("hidden");
  if (elements.localFallbackMessage) {
    elements.localFallbackMessage.textContent = message || "";
  }
}

function hideLocalFallback() {
  if (!elements.localFallback) return;
  elements.localFallback.classList.add("hidden");
}

async function handleFileSelection(event) {
  hideLocalFallback();
  const files = Array.from(event.target.files || []);
  const jsonFiles = files.filter((file) => file.name.toLowerCase().endsWith(".json"));

  if (!jsonFiles.length) {
    showLocalFallback("No JSON files found in that folder.");
    return;
  }

  const results = await Promise.allSettled(jsonFiles.map((file) => readJsonFile(file)));
  const sources = [];
  const cacheWrites = [];

  results.forEach((result) => {
    if (result.status !== "fulfilled") return;
    const file = result.value.file;
    const data = result.value.data;
    const match = SOURCE_FILES.find(
      (source) => source.url.toLowerCase() === file.name.toLowerCase()
    );
    const id = match ? match.id : file.name.replace(/\.json$/i, "");
    const label = match ? match.label : file.name.replace(/\.json$/i, "");
    const normalized = normalizeSource(id, data, label);
    pushSources(sources, normalized);
    cacheWrites.push(cacheSource(id, label, data));
  });

  if (!sources.length) {
    showLocalFallback("Could not read any JSON files. Check that the files are valid.");
    return;
  }

  if (cacheWrites.length) {
    await Promise.allSettled(cacheWrites);
  }

  applySources(sources);
}

async function readJsonFile(file) {
  const text = await file.text();
  const data = JSON.parse(text);
  return { file, data };
}

function slugify(text) {
  return String(text || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function buildDailyQuoteIndex() {
  const source = state.sourceById.get("daily");
  if (!source || !source.sections.length) {
    state.dailyQuotes = [];
    return;
  }

  const seen = new Set();
  state.dailyQuotes = source.sections
    .map((section) => {
      const meta = section.meta || {};
      const quote = meta.quote ? String(meta.quote).trim() : "";
      if (!quote || seen.has(quote)) return null;
      seen.add(quote);
      return {
        text: quote,
        date: meta.date || ""
      };
    })
    .filter(Boolean);
}

function openCache() {
  if (!("indexedDB" in window)) return Promise.resolve(null);
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(CACHE_DB_NAME, 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(CACHE_STORE)) {
        db.createObjectStore(CACHE_STORE, { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function getCachedSources() {
  try {
    const db = await openCache();
    if (!db) return [];
    return await new Promise((resolve) => {
      const tx = db.transaction(CACHE_STORE, "readonly");
      const store = tx.objectStore(CACHE_STORE);
      const request = store.getAll();
      request.onsuccess = () => resolve(request.result || []);
      request.onerror = () => resolve([]);
      tx.oncomplete = () => db.close();
      tx.onerror = () => db.close();
    });
  } catch (error) {
    return [];
  }
}

async function cacheSource(id, label, data) {
  try {
    const db = await openCache();
    if (!db) return;
    await new Promise((resolve) => {
      const tx = db.transaction(CACHE_STORE, "readwrite");
      const store = tx.objectStore(CACHE_STORE);
      store.put({
        id,
        label,
        data,
        updatedAt: Date.now()
      });
      tx.oncomplete = () => {
        db.close();
        resolve();
      };
      tx.onerror = () => {
        db.close();
        resolve();
      };
    });
  } catch (error) {
    return;
  }
}
