const SOURCE_FILES = [
  { id: "bbook", url: "bbook.json", label: "Big Book" },
  { id: "dictionary", url: "big_dictionary.json", label: "Dictionary" },
  { id: "twlvxtwlv", url: "twlvxtwlv.json", label: "Twelve Steps and Twelve Traditions" },
  { id: "daily", url: "daily.json", label: "Daily Reflections" }
];

const SOURCE_ORDER = ["steps", "bbook", "traditions", "dictionary", "daily"];

const SOURCE_HINTS = {
  steps: "Upload twlvxtwlv.json to see the Steps.",
  bbook: "Upload bbook.json to see the Big Book.",
  dictionary: "Upload big_dictionary.json to see the Dictionary.",
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

const WEEKDAY_LABELS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

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

const DAILY_NOTES_KEY = "daily-reflections-notes-v1";
const DAILY_FONT_SCALE_KEY = "daily-reflections-font-scale-v1";
const DAILY_FONT_SCALE_MIN = 0.85;
const DAILY_FONT_SCALE_MAX = 1.25;
const DAILY_FONT_SCALE_STEP = 0.05;
const DAILY_FALLBACK_YEAR = 2024;

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
  dailyQuotes: [],
  dailyNotes: {},
  dailyFontScale: 1,
  dailyCalendarMonthIndex: null,
  dictionaryIndex: new Map(),
  dictionaryRegex: null,
  dictionaryListScrollTop: null
};

const elements = {};

document.addEventListener("DOMContentLoaded", () => {
  cacheElements();
  initDefinitionTooltip();
  initDailyMenu();
  initDailyCalendarModal();
  bindEvents();
  loadDailyNotes();
  loadDailyFontScale();
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
    dictionary: document.getElementById("navDictionary"),
    daily: document.getElementById("navDaily")
  };
  elements.dropdownPanels = {
    steps: document.getElementById("dropdownSteps"),
    bbook: document.getElementById("dropdownBigBook"),
    traditions: document.getElementById("dropdownTraditions"),
    dictionary: document.getElementById("dropdownDictionary"),
    daily: document.getElementById("dropdownDaily")
  };
  elements.dropdownLists = {
    steps: document.getElementById("stepsList"),
    bbook: document.getElementById("bbookList"),
    traditions: document.getElementById("traditionsList"),
    dictionary: document.getElementById("dictionaryList")
  };
  elements.dailyMonthStrip = document.getElementById("dailyMonthStrip");
  elements.dailyDayList = document.getElementById("dailyDayList");
}

function initDefinitionTooltip() {
  const tooltip = document.createElement("div");
  tooltip.id = "definitionTooltip";
  tooltip.className = "definition-tooltip hidden";
  tooltip.setAttribute("role", "dialog");
  tooltip.setAttribute("aria-live", "polite");

  const term = document.createElement("div");
  term.className = "definition-term";
  const meta = document.createElement("div");
  meta.className = "definition-meta";
  const pron = document.createElement("span");
  pron.className = "definition-pronunciation";
  const pages = document.createElement("span");
  pages.className = "definition-pages";
  meta.appendChild(pron);
  meta.appendChild(pages);

  const parts = document.createElement("div");
  parts.className = "definition-parts";

  tooltip.appendChild(term);
  tooltip.appendChild(meta);
  tooltip.appendChild(parts);
  document.body.appendChild(tooltip);

  elements.definitionTooltip = tooltip;
  elements.definitionTooltipTerm = term;
  elements.definitionTooltipMeta = meta;
  elements.definitionTooltipPron = pron;
  elements.definitionTooltipPages = pages;
  elements.definitionTooltipParts = parts;
}

function initDailyMenu() {
  if (document.getElementById("dailyMenuModal")) return;
  const modal = document.createElement("div");
  modal.id = "dailyMenuModal";
  modal.className = "daily-modal";
  modal.setAttribute("aria-hidden", "true");

  modal.innerHTML = `
    <div class="daily-modal-backdrop" data-close="daily-menu"></div>
    <div class="daily-modal-card daily-menu-card" role="dialog" aria-modal="true" aria-labelledby="dailyMenuTitle">
      <div class="daily-modal-header">
        <div class="daily-label" id="dailyMenuTitle">Menu</div>
        <button class="daily-ghost daily-small" id="dailyMenuClose" type="button">Close</button>
      </div>
      <div class="daily-menu-layout">
        <section class="daily-card daily-menu-section daily-notes-panel">
          <div class="daily-card-header">
            <div class="daily-label">My note for this day</div>
          </div>
          <textarea id="dailyNoteField" rows="5" placeholder="Capture what resonates. Saved on this device."></textarea>
          <div class="daily-note-actions">
            <button id="dailySaveNote" type="button">Save note</button>
            <button class="daily-ghost" id="dailyDeleteNote" type="button">Delete note</button>
            <span class="daily-note-status" id="dailyNoteStatus" role="status" aria-live="polite"></span>
          </div>
          <div class="daily-note-review">
            <div class="daily-label">Review</div>
            <p class="daily-note-preview" id="dailyNotePreview">No saved note yet.</p>
          </div>
          <div class="daily-notes-library">
            <div class="daily-label">Saved notes</div>
            <div class="daily-notes-list" id="dailyNotesList"></div>
          </div>
        </section>
        <section class="daily-card daily-menu-section daily-future-panel">
          <div class="daily-label">Future space</div>
          <p class="daily-future-copy">
            Reserve this spot for streaks, reminders, or saved favorites.
          </p>
          <div class="daily-font-scale">
            <div class="daily-label">Font size</div>
            <div class="daily-font-scale-controls" role="group" aria-label="Font size controls">
              <button class="daily-ghost daily-small" id="dailyFontScaleDown" type="button">A-</button>
              <div class="daily-font-scale-value" id="dailyFontScaleValue">100%</div>
              <button class="daily-ghost daily-small" id="dailyFontScaleUp" type="button">A+</button>
            </div>
          </div>
        </section>
      </div>
    </div>
  `;

  document.body.appendChild(modal);

  elements.dailyMenuModal = modal;
  elements.dailyMenuClose = modal.querySelector("#dailyMenuClose");
  elements.dailyNoteField = modal.querySelector("#dailyNoteField");
  elements.dailySaveNote = modal.querySelector("#dailySaveNote");
  elements.dailyDeleteNote = modal.querySelector("#dailyDeleteNote");
  elements.dailyNoteStatus = modal.querySelector("#dailyNoteStatus");
  elements.dailyNotePreview = modal.querySelector("#dailyNotePreview");
  elements.dailyNotesList = modal.querySelector("#dailyNotesList");
  elements.dailyFontScaleDown = modal.querySelector("#dailyFontScaleDown");
  elements.dailyFontScaleUp = modal.querySelector("#dailyFontScaleUp");
  elements.dailyFontScaleValue = modal.querySelector("#dailyFontScaleValue");

  if (elements.dailyMenuClose) {
    elements.dailyMenuClose.addEventListener("click", closeDailyMenu);
  }
  modal.addEventListener("click", (event) => {
    if (event.target?.dataset?.close === "daily-menu") {
      closeDailyMenu();
    }
  });

  if (elements.dailySaveNote) {
    elements.dailySaveNote.addEventListener("click", saveDailyNote);
  }
  if (elements.dailyDeleteNote) {
    elements.dailyDeleteNote.addEventListener("click", deleteDailyNote);
  }
  if (elements.dailyFontScaleDown) {
    elements.dailyFontScaleDown.addEventListener("click", () => nudgeDailyFontScale(-1));
  }
  if (elements.dailyFontScaleUp) {
    elements.dailyFontScaleUp.addEventListener("click", () => nudgeDailyFontScale(1));
  }
}

function initDailyCalendarModal() {
  if (document.getElementById("dailyCalendarModal")) return;
  const modal = document.createElement("div");
  modal.id = "dailyCalendarModal";
  modal.className = "daily-modal";
  modal.setAttribute("aria-hidden", "true");

  modal.innerHTML = `
    <div class="daily-modal-backdrop" data-close="daily-calendar"></div>
    <div class="daily-modal-card daily-calendar-card" role="dialog" aria-modal="true" aria-labelledby="dailyCalendarTitle">
      <div class="daily-modal-header">
        <div class="daily-label" id="dailyCalendarTitle">Choose a day</div>
        <button class="daily-ghost daily-small" id="dailyCalendarClose" type="button">Close</button>
      </div>
      <div class="daily-calendar-controls">
        <button class="daily-ghost daily-small" id="dailyCalendarPrev" type="button">Prev month</button>
        <div class="daily-calendar-month" id="dailyCalendarMonthLabel">Month</div>
        <button class="daily-ghost daily-small" id="dailyCalendarNext" type="button">Next month</button>
      </div>
      <div class="daily-calendar-grid" id="dailyCalendarGrid" role="grid" aria-label="Monthly calendar"></div>
    </div>
  `;

  document.body.appendChild(modal);

  elements.dailyCalendarModal = modal;
  elements.dailyCalendarClose = modal.querySelector("#dailyCalendarClose");
  elements.dailyCalendarPrev = modal.querySelector("#dailyCalendarPrev");
  elements.dailyCalendarNext = modal.querySelector("#dailyCalendarNext");
  elements.dailyCalendarMonthLabel = modal.querySelector("#dailyCalendarMonthLabel");
  elements.dailyCalendarGrid = modal.querySelector("#dailyCalendarGrid");

  if (elements.dailyCalendarClose) {
    elements.dailyCalendarClose.addEventListener("click", closeDailyCalendar);
  }
  if (elements.dailyCalendarPrev) {
    elements.dailyCalendarPrev.addEventListener("click", () => adjustDailyCalendarMonth(-1));
  }
  if (elements.dailyCalendarNext) {
    elements.dailyCalendarNext.addEventListener("click", () => adjustDailyCalendarMonth(1));
  }

  modal.addEventListener("click", (event) => {
    if (event.target?.dataset?.close === "daily-calendar") {
      closeDailyCalendar();
    }
  });
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

  document.addEventListener("click", (event) => {
    const target = event.target;
    const mark = target?.closest?.("mark.dictionary-term");
    if (mark) {
      showDefinitionTooltip(mark);
      return;
    }
    if (elements.definitionTooltip && elements.definitionTooltip.contains(target)) return;
    hideDefinitionTooltip();
  });

  window.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      closeDropdown();
      closeDailyMenu();
      closeDailyCalendar();
      hideDefinitionTooltip();
    }
  });

  document.addEventListener("keydown", (event) => {
    if (event.key !== "Enter" && event.key !== " ") return;
    const target = event.target;
    if (!target || !target.classList?.contains("dictionary-term")) return;
    event.preventDefault();
    showDefinitionTooltip(target);
  });

  window.addEventListener("scroll", hideDefinitionTooltip, { passive: true });
  window.addEventListener("resize", hideDefinitionTooltip);
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
    const meta = {
      type: section.type || "section",
      number: section.number || null,
      subtitle: section.subtitle || ""
    };

    if (section.pronunciation || section.pages || section.definitionParts) {
      meta.pronunciation = section.pronunciation || "";
      meta.pages = section.pages || "";
      meta.definitionParts = Array.isArray(section.definitionParts)
        ? section.definitionParts
        : [];
    }

    return {
      id: section.id || `${id}-section-${index + 1}`,
      heading,
      paragraphs: content.filter(Boolean),
      meta
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
  buildDictionaryIndex();

  const hasDaily = state.sourceById.has("daily");
  if (!state.activeSourceId || !state.sourceById.has(state.activeSourceId)) {
    state.activeSourceId = hasDaily ? "daily" : getFirstAvailableSourceId();
  }

  state.activeSectionBySource.forEach((_value, key) => {
    if (!state.sourceById.has(key)) {
      state.activeSectionBySource.delete(key);
    }
  });

  if (hasDaily && !state.activeSectionBySource.has("daily")) {
    const dailySource = state.sourceById.get("daily");
    const todayId = dailySource ? getDefaultDailyId(dailySource) : null;
    if (todayId) {
      state.activeSectionBySource.set("daily", todayId);
    }
  }

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
  } else if (source.id === "dictionary") {
    renderDictionarySection(source, section);
  } else {
    closeDailyMenu();
    closeDailyCalendar();
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

function renderDictionarySection(source, section) {
  elements.activeKicker.textContent = source.title;

  const activeSectionId = ensureActiveSection(source);
  const activeSection = activeSectionId ? source.sectionById.get(activeSectionId) : section;

  elements.activeTitle.textContent = activeSection?.heading || source.title;
  elements.activeMeta.textContent = source.author || "";
  elements.content.innerHTML = "";

  if (!source.sections.length) {
    const empty = document.createElement("div");
    empty.className = "panel-note";
    empty.textContent = "No content available.";
    elements.content.appendChild(empty);
    return;
  }

  const layout = document.createElement("div");
  layout.className = "dictionary-layout";

  const listPane = document.createElement("aside");
  listPane.className = "dictionary-pane dictionary-pane-list";

  const listHeader = document.createElement("div");
  listHeader.className = "dictionary-list-header";

  const listTitle = document.createElement("div");
  listTitle.className = "dictionary-list-title";
  listTitle.textContent = source.title || "Dictionary";

  const listMeta = document.createElement("div");
  listMeta.className = "dictionary-list-meta";
  listMeta.textContent = `${source.sections.length} terms`;

  listHeader.appendChild(listTitle);
  listHeader.appendChild(listMeta);
  listPane.appendChild(listHeader);

  const listScroll = document.createElement("div");
  listScroll.className = "dictionary-list-scroll";
  listScroll.setAttribute("role", "listbox");
  listScroll.setAttribute("aria-label", "Dictionary term list");

  const activeId = activeSection?.id || activeSectionId;
  const letterTargets = new Map();
  const listItems = [];

  source.sections.forEach((entry, index) => {
    const termText = entry.heading || "Untitled";
    const definitionParts = getDefinitionPartsFromSection(entry);
    const definitionText = definitionParts.length
      ? definitionParts.join(" / ")
      : (entry.paragraphs || []).filter(Boolean).join(" ").trim();
    const button = document.createElement("button");
    button.type = "button";
    button.className = "dictionary-list-item";
    button.setAttribute(
      "aria-label",
      definitionText ? `${termText}. ${definitionText}` : termText
    );
    button.style.setProperty("--i", index);
    button.setAttribute("role", "option");
    button.setAttribute("aria-selected", entry.id === activeId ? "true" : "false");
    const letter = getDictionaryIndexLetter(termText);
    button.dataset.letter = letter;
    const termSpan = document.createElement("span");
    termSpan.className = "dictionary-list-term";
    termSpan.textContent = termText;
    button.appendChild(termSpan);
    if (definitionText) {
      const definitionSpan = document.createElement("span");
      definitionSpan.className = "dictionary-list-definition";
      definitionSpan.textContent = definitionText;
      button.appendChild(definitionSpan);
    }
    if (entry.id === activeId) {
      button.classList.add("active");
    }
    button.addEventListener("click", () => {
      if (listScroll.dataset.dragging === "true") return;
      state.dictionaryListScrollTop = listScroll.scrollTop;
      state.activeSectionBySource.set(source.id, entry.id);
      setActiveSource(source.id);
      closeDropdown();
    });
    listScroll.appendChild(button);
    listItems.push(button);
    if (!letterTargets.has(letter)) {
      letterTargets.set(letter, button);
    }
  });

  const listBody = document.createElement("div");
  listBody.className = "dictionary-list-body";
  listBody.appendChild(listScroll);

  const alphabetNav = document.createElement("nav");
  alphabetNav.className = "dictionary-alpha";
  alphabetNav.setAttribute("aria-label", "Alphabet shortcuts");

  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");
  if (letterTargets.has("#")) {
    alphabet.unshift("#");
  }

  const alphaButtons = new Map();
  alphabet.forEach((letter) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "dictionary-alpha-item";
    button.textContent = letter;
    const target = letterTargets.get(letter);
    if (!target) {
      button.disabled = true;
    } else {
      button.addEventListener("click", () => {
        listScroll.scrollTo({ top: Math.max(0, target.offsetTop - 4), behavior: "smooth" });
      });
    }
    alphaButtons.set(letter, button);
    alphabetNav.appendChild(button);
  });

  listBody.appendChild(alphabetNav);
  listPane.appendChild(listBody);

  let letterRaf = null;
  const updateAlphabetActive = () => {
    letterRaf = null;
    if (!listItems.length) return;
    const scrollTop = listScroll.scrollTop;
    const current =
      listItems.find((item) => item.offsetTop + item.offsetHeight > scrollTop + 2) ||
      listItems[0];
    const letter = current?.dataset?.letter || "";
    alphaButtons.forEach((button, key) => {
      button.classList.toggle("active", key === letter);
    });
  };

  listScroll.addEventListener("scroll", () => {
    state.dictionaryListScrollTop = listScroll.scrollTop;
    if (letterRaf) return;
    letterRaf = requestAnimationFrame(updateAlphabetActive);
  });

  elements.content.appendChild(layout);

  const entryPane = document.createElement("section");
  entryPane.className = "dictionary-pane dictionary-pane-entry";
  entryPane.appendChild(buildDictionaryEntryCard(source, activeSection));

  layout.appendChild(listPane);
  layout.appendChild(entryPane);

  attachDictionaryScroller(listScroll);

  if (typeof state.dictionaryListScrollTop === "number") {
    const maxScroll = Math.max(0, listScroll.scrollHeight - listScroll.clientHeight);
    listScroll.scrollTop = Math.min(state.dictionaryListScrollTop, maxScroll);
  } else {
    const activeButton = listScroll.querySelector(".dictionary-list-item.active");
    if (activeButton) {
      activeButton.scrollIntoView({ block: "center" });
    }
  }
  updateAlphabetActive();
}

function buildDictionaryEntryCard(source, section) {
  const wrapper = document.createElement("section");
  wrapper.className = "section-block dictionary-entry";

  if (!section) {
    const empty = document.createElement("div");
    empty.className = "panel-note";
    empty.textContent = "No content available.";
    wrapper.appendChild(empty);
    return wrapper;
  }

  const header = document.createElement("header");
  header.className = "dictionary-entry-header";

  const term = document.createElement("h2");
  term.className = "dictionary-entry-term";
  term.textContent = section.heading;
  header.appendChild(term);

  const metaRow = document.createElement("div");
  metaRow.className = "dictionary-entry-meta";

  const pronunciation = section.meta?.pronunciation;
  if (pronunciation) {
    const pron = document.createElement("span");
    pron.className = "dictionary-entry-pronunciation";
    pron.textContent = `(${pronunciation})`;
    metaRow.appendChild(pron);
  }

  const pages = section.meta?.pages;
  if (pages) {
    const page = document.createElement("span");
    page.className = "dictionary-entry-pages";
    page.textContent = pages;
    metaRow.appendChild(page);
  }

  if (metaRow.childNodes.length) {
    header.appendChild(metaRow);
  }

  wrapper.appendChild(header);

  const parts = getDefinitionPartsFromSection(section);
  if (parts.length) {
    const definitions = document.createElement("div");
    definitions.className = "dictionary-entry-definitions";
    definitions.id = buildParagraphId(source.key, section.index, 0);
    renderDefinitionParts(definitions, parts, {
      partClass: "dictionary-entry-part",
      highlight: true
    });
    wrapper.appendChild(definitions);
  } else {
    section.paragraphs.forEach((text, index) => {
      const paragraph = document.createElement("p");
      paragraph.className = "para";
      paragraph.id = buildParagraphId(source.key, section.index, index);
      appendHighlightedText(paragraph, text, state.searchTerm, { includeCrossrefs: true });
      wrapper.appendChild(paragraph);
    });
  }

  return wrapper;
}

function getDictionaryIndexLetter(value) {
  const match = String(value || "").trim().match(/[A-Za-z]/);
  return match ? match[0].toUpperCase() : "#";
}

function attachDictionaryScroller(container) {
  if (!container || container.dataset.scrollerAttached === "true") return;
  container.dataset.scrollerAttached = "true";
  container.dataset.dragging = "false";
  container.style.touchAction = "none";

  let isDown = false;
  let startY = 0;
  let startScroll = 0;
  let lastY = 0;
  let lastTime = 0;
  let velocity = 0;
  let rafId = null;

  const stopMomentum = () => {
    if (rafId) {
      cancelAnimationFrame(rafId);
      rafId = null;
    }
  };

  const onPointerDown = (event) => {
    if (event.pointerType === "mouse" && event.button !== 0) return;
    stopMomentum();
    container.dataset.dragging = "false";
    isDown = true;
    startY = event.clientY;
    lastY = event.clientY;
    startScroll = container.scrollTop;
    lastTime = performance.now();
    velocity = 0;
    container.setPointerCapture(event.pointerId);
  };

  const onPointerMove = (event) => {
    if (!isDown) return;
    const currentY = event.clientY;
    const now = performance.now();
    const delta = currentY - startY;
    if (Math.abs(delta) > 6) {
      container.dataset.dragging = "true";
    }
    container.scrollTop = startScroll - delta;
    const dt = now - lastTime;
    if (dt > 0) {
      velocity = (currentY - lastY) / dt;
    }
    lastY = currentY;
    lastTime = now;
  };

  const onPointerUp = (event) => {
    if (!isDown) return;
    isDown = false;
    container.releasePointerCapture(event.pointerId);

    const absVelocity = Math.abs(velocity);
    if (absVelocity > 0.05) {
      let currentVelocity = velocity * 18;
      const decay = 0.92;
      const maxScroll = Math.max(0, container.scrollHeight - container.clientHeight);

      const step = () => {
        if (Math.abs(currentVelocity) < 0.3) {
          rafId = null;
          return;
        }

        const next = container.scrollTop - currentVelocity;
        if (next <= 0) {
          container.scrollTop = 0;
          rafId = null;
          return;
        }
        if (next >= maxScroll) {
          container.scrollTop = maxScroll;
          rafId = null;
          return;
        }

        container.scrollTop = next;
        currentVelocity *= decay;
        rafId = requestAnimationFrame(step);
      };

      rafId = requestAnimationFrame(step);
    }

    window.setTimeout(() => {
      container.dataset.dragging = "false";
    }, 0);
  };

  container.addEventListener("pointerdown", onPointerDown);
  container.addEventListener("pointermove", onPointerMove);
  container.addEventListener("pointerup", onPointerUp);
  container.addEventListener("pointercancel", onPointerUp);
  container.addEventListener("pointerleave", onPointerUp);
  container.addEventListener("wheel", stopMomentum, { passive: true });
}

function renderDailySection(source, section) {
  if (!section) {
    closeDailyMenu();
    closeDailyCalendar();
    renderBookSection(source, section);
    return;
  }

  const meta = section.meta || {};
  if (meta.month) {
    state.activeMonth = meta.month;
  }

  elements.activeKicker.textContent = "";
  elements.activeTitle.textContent = "Daily Reflections";
  elements.activeMeta.textContent = meta.date || "";

  elements.content.innerHTML = "";

  const shell = document.createElement("section");
  shell.className = "daily-shell";
  shell.style.setProperty("--type-scale", (state.dailyFontScale || 1).toString());

  const topbar = document.createElement("header");
  topbar.className = "topbar";

  const topbarSpacer = document.createElement("div");
  topbarSpacer.className = "topbar-spacer";

  const topbarTitle = document.createElement("h1");
  topbarTitle.className = "topbar-title";
  topbarTitle.textContent = "Daily Reflections";

  const menuButton = document.createElement("button");
  menuButton.type = "button";
  menuButton.className = "icon-button daily-menu-button";
  menuButton.setAttribute("aria-label", "Open menu");
  menuButton.setAttribute("aria-controls", "dailyMenuModal");
  menuButton.setAttribute("aria-expanded", "false");
  const hamburger = document.createElement("span");
  hamburger.className = "hamburger";
  menuButton.appendChild(hamburger);

  topbar.appendChild(topbarSpacer);
  topbar.appendChild(topbarTitle);
  topbar.appendChild(menuButton);
  shell.appendChild(topbar);
  attachDailyMenuButton(menuButton);

  const card = document.createElement("section");
  card.className = "card current";
  card.setAttribute("aria-live", "polite");

  const header = document.createElement("div");
  header.className = "card-header";

  const monthLabel = document.createElement("div");
  monthLabel.className = "label current-month";
  monthLabel.textContent = meta.month || MONTHS[new Date().getMonth()];

  const dayWrap = document.createElement("div");
  dayWrap.className = "day-wrap";
  const dayLabel = document.createElement("span");
  dayLabel.className = "day";
  const dayNumber = Number(meta.day) || new Date().getDate();
  dayLabel.textContent = String(dayNumber).padStart(2, "0");
  dayWrap.appendChild(dayLabel);

  header.appendChild(monthLabel);
  header.appendChild(dayWrap);
  card.appendChild(header);

  const nav = document.createElement("div");
  nav.className = "entry-nav";
  nav.setAttribute("aria-label", "Entry navigation");

  const makeIconButton = (label, path) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "ghost small icon-only";
    button.setAttribute("aria-label", label);
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", "0 0 24 24");
    svg.setAttribute("aria-hidden", "true");
    svg.classList.add("nav-icon");
    const svgPath = document.createElementNS("http://www.w3.org/2000/svg", "path");
    svgPath.setAttribute("d", path);
    svg.appendChild(svgPath);
    button.appendChild(svg);
    return button;
  };

  const prevButton = makeIconButton("Previous day", "M15 6l-6 6 6 6");
  const todayButton = document.createElement("button");
  todayButton.type = "button";
  todayButton.className = "ghost small daily-today";
  todayButton.textContent = "Today";
  const nextButton = makeIconButton("Next day", "M9 6l6 6-6 6");
  const randomButton = document.createElement("button");
  randomButton.type = "button";
  randomButton.className = "ghost small";
  randomButton.textContent = "God's Pick";
  const calendarButton = makeIconButton("Open calendar", "M7 3v4M17 3v4M4 8h16M6 6h12a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2z");

  prevButton.addEventListener("click", () => moveDailyDay(-1));
  todayButton.addEventListener("click", showTodayDaily);
  nextButton.addEventListener("click", () => moveDailyDay(1));
  randomButton.addEventListener("click", showRandomDaily);
  calendarButton.addEventListener("click", openDailyPicker);

  nav.appendChild(prevButton);
  nav.appendChild(todayButton);
  nav.appendChild(nextButton);
  nav.appendChild(randomButton);
  nav.appendChild(calendarButton);
  card.appendChild(nav);

  const title = document.createElement("h2");
  title.textContent = meta.title || source.title;
  card.appendChild(title);

  let paragraphIndex = 0;
  if (meta.quote) {
    const quote = document.createElement("p");
    quote.className = "quote";
    quote.id = buildParagraphId(source.key, section.index, paragraphIndex);
    appendHighlightedText(quote, meta.quote, state.searchTerm, { includeCrossrefs: false });
    card.appendChild(quote);
    paragraphIndex += 1;
  }

  if (meta.source) {
    const sourceLine = document.createElement("p");
    sourceLine.className = "source";
    sourceLine.textContent = meta.source;
    card.appendChild(sourceLine);
  }

  if (meta.reflection) {
    const reflection = document.createElement("p");
    reflection.className = "reflection";
    reflection.id = buildParagraphId(source.key, section.index, paragraphIndex);
    appendHighlightedText(reflection, meta.reflection, state.searchTerm, {
      includeCrossrefs: false
    });
    card.appendChild(reflection);
  }

  const tags = document.createElement("div");
  tags.className = "tags";

  const pageTag = document.createElement("span");
  pageTag.className = "tag tag-page";
  pageTag.textContent = meta.pageIndex ? `Page ${meta.pageIndex}` : "Daily";

  const shareButton = document.createElement("button");
  shareButton.type = "button";
  shareButton.className = "ghost small share-pill";
  shareButton.textContent = "Share";
  shareButton.addEventListener("click", () => shareDailyEntry(section));

  tags.appendChild(pageTag);
  tags.appendChild(shareButton);
  card.appendChild(tags);

  shell.appendChild(card);
  elements.content.appendChild(shell);
  syncDailyNoteField();
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
  renderTocDropdown("dictionary");
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
    if (match.entry) {
      mark.dataset.term = match.term || safeText.slice(match.start, match.end);
      mark.setAttribute("role", "button");
      mark.setAttribute("tabindex", "0");
      mark.setAttribute("aria-haspopup", "dialog");
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

  collectDictionaryMatches(text, candidates);

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

function collectDictionaryMatches(text, output) {
  if (!state.dictionaryRegex || !state.dictionaryIndex.size) return;
  const regex = state.dictionaryRegex;
  regex.lastIndex = 0;
  let match;
  while ((match = regex.exec(text)) !== null) {
    if (!match[0]) continue;
    const term = match[0];
    const entry = state.dictionaryIndex.get(term.toLowerCase());
    if (!entry) continue;
    output.push({
      start: match.index,
      end: match.index + term.length,
      className: "dictionary-term",
      priority: 0,
      entry,
      term
    });
    if (match.index === regex.lastIndex) {
      regex.lastIndex += 1;
    }
  }
}

function buildCaseInsensitiveNeedle(text) {
  const trimmed = String(text || "").trim();
  if (!trimmed) return "";
  return trimmed.toLowerCase();
}

function splitDefinitionParts(text) {
  return String(text || "")
    .split(/\s*\/{1,2}\s*/)
    .map((part) => part.trim())
    .filter(Boolean);
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

function attachDailyMenuButton(button) {
  if (!button) return;
  elements.dailyMenuButton = button;
  const menuOpen = isDailyMenuOpen();
  button.setAttribute("aria-expanded", menuOpen ? "true" : "false");
  button.setAttribute("aria-label", menuOpen ? "Close menu" : "Open menu");
  button.addEventListener("click", () => {
    if (isDailyMenuOpen()) {
      closeDailyMenu();
    } else {
      openDailyMenu();
    }
  });
}

function isDailyMenuOpen() {
  return Boolean(elements.dailyMenuModal?.classList.contains("open"));
}

function isDailyCalendarOpen() {
  return Boolean(elements.dailyCalendarModal?.classList.contains("open"));
}

function openDailyMenu() {
  if (!elements.dailyMenuModal) return;
  elements.dailyMenuModal.classList.add("open");
  elements.dailyMenuModal.setAttribute("aria-hidden", "false");
  if (elements.dailyMenuButton) {
    elements.dailyMenuButton.setAttribute("aria-expanded", "true");
    elements.dailyMenuButton.setAttribute("aria-label", "Close menu");
  }
  document.body.classList.add("daily-menu-open");
  syncDailyNoteField();
  renderDailyNotesList();
  elements.dailyNoteField?.focus();
}

function closeDailyMenu() {
  if (!elements.dailyMenuModal) return;
  elements.dailyMenuModal.classList.remove("open");
  elements.dailyMenuModal.setAttribute("aria-hidden", "true");
  if (elements.dailyMenuButton) {
    elements.dailyMenuButton.setAttribute("aria-expanded", "false");
    elements.dailyMenuButton.setAttribute("aria-label", "Open menu");
  }
  document.body.classList.remove("daily-menu-open");
}

function openDailyCalendar() {
  if (!elements.dailyCalendarModal) return;
  if (!getDailySource()) return;
  const active = getActiveDailySection();
  const activeMonthIndex = active?.meta?.month
    ? MONTHS.findIndex(
        (month) => month.toLowerCase() === String(active.meta.month).toLowerCase()
      )
    : new Date().getMonth();
  state.dailyCalendarMonthIndex =
    activeMonthIndex >= 0 ? activeMonthIndex : new Date().getMonth();
  closeDropdown();
  renderDailyCalendar();
  elements.dailyCalendarModal.classList.add("open");
  elements.dailyCalendarModal.setAttribute("aria-hidden", "false");
}

function closeDailyCalendar() {
  if (!elements.dailyCalendarModal) return;
  elements.dailyCalendarModal.classList.remove("open");
  elements.dailyCalendarModal.setAttribute("aria-hidden", "true");
}

function adjustDailyCalendarMonth(direction) {
  const currentIndex = state.dailyCalendarMonthIndex ?? new Date().getMonth();
  state.dailyCalendarMonthIndex = (currentIndex + direction + 12) % 12;
  renderDailyCalendar();
}

function renderDailyCalendar() {
  if (!elements.dailyCalendarGrid || !elements.dailyCalendarMonthLabel) return;
  const monthIndex = state.dailyCalendarMonthIndex ?? new Date().getMonth();
  const firstDay = new Date(DAILY_FALLBACK_YEAR, monthIndex, 1).getDay();
  const days = daysInMonth(monthIndex);
  const today = new Date();

  elements.dailyCalendarMonthLabel.textContent = MONTHS[monthIndex];
  elements.dailyCalendarGrid.innerHTML = "";

  WEEKDAY_LABELS.forEach((label) => {
    const cell = document.createElement("div");
    cell.className = "daily-calendar-label";
    cell.setAttribute("role", "columnheader");
    cell.textContent = label;
    elements.dailyCalendarGrid.appendChild(cell);
  });

  for (let i = 0; i < firstDay; i += 1) {
    const empty = document.createElement("div");
    empty.className = "daily-calendar-empty";
    empty.setAttribute("aria-hidden", "true");
    elements.dailyCalendarGrid.appendChild(empty);
  }

  for (let day = 1; day <= days; day += 1) {
    const section = getDailySectionByDate(monthIndex, day);
    const button = document.createElement("button");
    button.type = "button";
    button.className = "daily-calendar-day";
    if (!section) {
      button.disabled = true;
    }

    if (today.getMonth() === monthIndex && today.getDate() === day) {
      button.classList.add("today");
    }

    const active = getActiveDailySection();
    if (
      active?.meta?.month &&
      active?.meta?.day &&
      MONTHS[monthIndex].toLowerCase() === String(active.meta.month).toLowerCase() &&
      Number(active.meta.day) === day
    ) {
      button.classList.add("selected");
    }

    button.textContent = String(day);
    if (section) {
      button.addEventListener("click", () => {
        jumpToDailyDate(monthIndex, day);
        closeDailyCalendar();
      });
    }
    elements.dailyCalendarGrid.appendChild(button);
  }
}

function loadDailyNotes() {
  try {
    state.dailyNotes = JSON.parse(localStorage.getItem(DAILY_NOTES_KEY) || "{}");
  } catch (err) {
    state.dailyNotes = {};
  }
  renderDailyNotesList();
}

function saveDailyNotes() {
  try {
    localStorage.setItem(DAILY_NOTES_KEY, JSON.stringify(state.dailyNotes || {}));
  } catch (err) {
    // Ignore storage failures.
  }
}

function currentDailyNoteKey() {
  const section = getActiveDailySection();
  const meta = section?.meta || {};
  if (!meta.month || !meta.day) return null;
  return `${meta.month}-${meta.day}`;
}

function syncDailyNoteField() {
  if (!elements.dailyNoteField) return;
  const key = currentDailyNoteKey();
  elements.dailyNoteField.value = (key && state.dailyNotes[key]) || "";
  if (elements.dailyNoteStatus) {
    elements.dailyNoteStatus.textContent = "";
  }
  syncDailyNotePreview();
}

function saveDailyNote() {
  const key = currentDailyNoteKey();
  if (!key || !elements.dailyNoteField) return;
  state.dailyNotes[key] = elements.dailyNoteField.value.trim();
  saveDailyNotes();
  if (elements.dailyNoteStatus) {
    elements.dailyNoteStatus.textContent = "Saved locally";
    setTimeout(() => {
      if (elements.dailyNoteStatus) {
        elements.dailyNoteStatus.textContent = "";
      }
    }, 1500);
  }
  syncDailyNotePreview();
  renderDailyNotesList();
}

function deleteDailyNote() {
  const key = currentDailyNoteKey();
  if (!key) return;
  if (!state.dailyNotes[key]) {
    if (elements.dailyNoteStatus) {
      elements.dailyNoteStatus.textContent = "No saved note to delete.";
      setTimeout(() => {
        if (elements.dailyNoteStatus) {
          elements.dailyNoteStatus.textContent = "";
        }
      }, 1500);
    }
    return;
  }
  delete state.dailyNotes[key];
  saveDailyNotes();
  syncDailyNoteField();
  renderDailyNotesList();
  if (elements.dailyNoteStatus) {
    elements.dailyNoteStatus.textContent = "Note deleted";
    setTimeout(() => {
      if (elements.dailyNoteStatus) {
        elements.dailyNoteStatus.textContent = "";
      }
    }, 1500);
  }
}

function syncDailyNotePreview() {
  if (!elements.dailyNotePreview) return;
  const key = currentDailyNoteKey();
  const savedNote = key ? state.dailyNotes[key] : "";
  elements.dailyNotePreview.textContent = savedNote || "No saved note yet.";
}

function renderDailyNotesList() {
  if (!elements.dailyNotesList) return;
  const noteKeys = Object.keys(state.dailyNotes || {});
  if (!noteKeys.length) {
    elements.dailyNotesList.innerHTML = '<div class="daily-empty">No saved notes yet.</div>';
    return;
  }

  const noteData = noteKeys
    .map((key) => {
      const [monthName, dayValue] = key.split("-");
      const day = Number(dayValue);
      const monthIndex = MONTHS.findIndex(
        (month) => month.toLowerCase() === String(monthName || "").toLowerCase()
      );
      const section = getDailySectionByDate(monthIndex, day);
      return {
        key,
        day,
        monthName,
        monthIndex,
        title: section?.meta?.title || "Daily Reflection",
        note: state.dailyNotes[key] || ""
      };
    })
    .filter((item) => item.monthIndex >= 0 && item.day)
    .sort((a, b) => {
      if (a.monthIndex !== b.monthIndex) return a.monthIndex - b.monthIndex;
      return a.day - b.day;
    });

  elements.dailyNotesList.innerHTML = "";
  noteData.forEach((item) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "daily-note-item";
    button.dataset.month = String(item.monthIndex);
    button.dataset.day = String(item.day);

    const title = document.createElement("div");
    title.className = "daily-note-item-title";
    title.textContent = item.title;

    const meta = document.createElement("div");
    meta.className = "daily-note-item-meta";
    meta.textContent = `${item.monthName} ${item.day}`;

    const snippet = document.createElement("div");
    snippet.className = "daily-note-item-snippet";
    snippet.textContent = item.note.trim() || "Empty note.";

    button.appendChild(title);
    button.appendChild(meta);
    button.appendChild(snippet);
    button.addEventListener("click", () => {
      jumpToDailyDate(item.monthIndex, item.day);
      closeDailyMenu();
    });

    elements.dailyNotesList.appendChild(button);
  });
}

function clampDailyFontScale(value) {
  return Math.min(DAILY_FONT_SCALE_MAX, Math.max(DAILY_FONT_SCALE_MIN, value));
}

function updateDailyFontScaleUI() {
  if (!elements.dailyFontScaleValue) return;
  const percent = Math.round((state.dailyFontScale || 1) * 100);
  elements.dailyFontScaleValue.textContent = `${percent}%`;
  if (elements.dailyFontScaleDown) {
    elements.dailyFontScaleDown.disabled =
      state.dailyFontScale <= DAILY_FONT_SCALE_MIN + 0.001;
  }
  if (elements.dailyFontScaleUp) {
    elements.dailyFontScaleUp.disabled =
      state.dailyFontScale >= DAILY_FONT_SCALE_MAX - 0.001;
  }
}

function setDailyFontScale(value, options = {}) {
  const nextValue = clampDailyFontScale(value);
  state.dailyFontScale = Number(nextValue.toFixed(2));
  const shell = document.querySelector(".daily-shell");
  if (shell) {
    shell.style.setProperty("--type-scale", state.dailyFontScale.toString());
  }
  updateDailyFontScaleUI();
  if (options.persist === false) return;
  try {
    localStorage.setItem(DAILY_FONT_SCALE_KEY, state.dailyFontScale.toString());
  } catch (err) {
    // Ignore storage failures.
  }
}

function loadDailyFontScale() {
  try {
    const stored = Number(localStorage.getItem(DAILY_FONT_SCALE_KEY));
    if (!Number.isNaN(stored) && stored > 0) {
      setDailyFontScale(stored, { persist: false });
      return;
    }
  } catch (err) {
    // Ignore storage failures.
  }
  setDailyFontScale(1, { persist: false });
}

function nudgeDailyFontScale(direction) {
  setDailyFontScale((state.dailyFontScale || 1) + DAILY_FONT_SCALE_STEP * direction);
}

function getDailySource() {
  return state.sourceById.get("daily") || null;
}

function getActiveDailySection() {
  const source = getDailySource();
  if (!source) return null;
  const sectionId = state.activeSectionBySource.get("daily");
  return sectionId ? source.sectionById.get(sectionId) : null;
}

function getDailySectionByDate(monthIndex, day) {
  const source = getDailySource();
  if (!source) return null;
  const monthName = MONTHS[monthIndex];
  if (!monthName || !day) return null;
  return (
    source.sections.find(
      (section) =>
        section.meta?.month === monthName && Number(section.meta?.day) === Number(day)
    ) || null
  );
}

function jumpToDailySection(section) {
  if (!section) return;
  state.activeSectionBySource.set("daily", section.id);
  setActiveSource("daily");
}

function jumpToDailyDate(monthIndex, day) {
  const section = getDailySectionByDate(monthIndex, day);
  if (section) {
    jumpToDailySection(section);
  }
}

function showTodayDaily() {
  const today = new Date();
  jumpToDailyDate(today.getMonth(), today.getDate());
}

function daysInMonth(monthIndex) {
  return new Date(DAILY_FALLBACK_YEAR, monthIndex + 1, 0).getDate();
}

function shiftDailyDay(monthIndex, day, direction) {
  let nextMonth = monthIndex;
  let nextDay = day + direction;
  if (direction > 0 && nextDay > daysInMonth(nextMonth)) {
    nextMonth = (nextMonth + 1) % 12;
    nextDay = 1;
  }
  if (direction < 0 && nextDay < 1) {
    nextMonth = (nextMonth + 11) % 12;
    nextDay = daysInMonth(nextMonth);
  }
  return { monthIndex: nextMonth, day: nextDay };
}

function moveDailyDay(direction) {
  const source = getDailySource();
  if (!source) return;
  const active = getActiveDailySection();
  let monthIndex = active?.meta?.month
    ? MONTHS.findIndex(
        (month) => month.toLowerCase() === String(active.meta.month).toLowerCase()
      )
    : new Date().getMonth();
  let day = Number(active?.meta?.day) || new Date().getDate();
  if (monthIndex < 0) {
    monthIndex = new Date().getMonth();
  }

  let safety = 0;
  ({ monthIndex, day } = shiftDailyDay(monthIndex, day, direction));
  while (safety < 370) {
    const section = getDailySectionByDate(monthIndex, day);
    if (section) {
      jumpToDailySection(section);
      return;
    }
    ({ monthIndex, day } = shiftDailyDay(monthIndex, day, direction));
    safety += 1;
  }
}

function showRandomDaily() {
  const source = getDailySource();
  if (!source || !source.sections.length) return;
  const pick = source.sections[Math.floor(Math.random() * source.sections.length)];
  if (pick) {
    jumpToDailySection(pick);
  }
}

function openDailyPicker() {
  openDailyCalendar();
}

async function shareDailyEntry(section) {
  if (!section) return;
  const meta = section.meta || {};
  const text = `${meta.title || section.heading}\n${meta.quote || ""}\n\n${meta.reflection || ""}`.trim();
  const sharePayload = {
    title: "Daily Reflections",
    text
  };

  if (navigator.share) {
    try {
      await navigator.share(sharePayload);
    } catch (err) {
      // user canceled share
    }
  } else if (navigator.clipboard) {
    try {
      await navigator.clipboard.writeText(text);
      if (elements.dailyNoteStatus) {
        elements.dailyNoteStatus.textContent = "Copied for sharing";
        setTimeout(() => {
          if (elements.dailyNoteStatus) {
            elements.dailyNoteStatus.textContent = "";
          }
        }, 1600);
      }
    } catch (err) {
      // Ignore clipboard errors.
    }
  }
}

function buildDictionaryIndex() {
  const source = state.sourceById.get("dictionary");
  if (!source || !source.sections.length) {
    state.dictionaryIndex = new Map();
    state.dictionaryRegex = null;
    return;
  }

  const index = new Map();
  const terms = [];
  source.sections.forEach((section) => {
    const term = String(section.heading || section.title || "").trim();
    if (term.length <= 1) return;
    const definition = (section.paragraphs || []).filter(Boolean).join(" ").trim();
    if (!definition) return;
    const parts = getDefinitionPartsFromSection(section);
    const key = term.toLowerCase();
    if (index.has(key)) return;
    index.set(key, {
      term,
      definition,
      parts,
      pronunciation: section.meta?.pronunciation || "",
      pages: section.meta?.pages || ""
    });
    terms.push(term);
  });

  state.dictionaryIndex = index;
  state.dictionaryRegex = buildDictionaryRegex(terms);
}

function buildDictionaryRegex(terms) {
  const filtered = (terms || []).map((term) => String(term || "").trim()).filter(Boolean);
  if (!filtered.length) return null;
  filtered.sort((a, b) => b.length - a.length);
  const pattern = filtered.map(escapeRegex).join("|");
  return new RegExp(`\\b(?:${pattern})\\b`, "gi");
}

function getDictionaryEntry(term) {
  const key = String(term || "").trim().toLowerCase();
  if (!key) return null;
  return state.dictionaryIndex.get(key) || null;
}

function getDefinitionPartsFromSection(section) {
  const metaParts = section.meta?.definitionParts;
  if (Array.isArray(metaParts) && metaParts.length) {
    return metaParts.map((part) => String(part || "").trim()).filter(Boolean);
  }
  const definition = (section.paragraphs || []).filter(Boolean).join(" ").trim();
  return splitDefinitionParts(definition);
}

function renderDefinitionParts(container, parts, options = {}) {
  container.innerHTML = "";
  const className = options.partClass || "definition-part";
  const highlight = options.highlight ? { includeCrossrefs: false } : null;
  parts.forEach((part) => {
    const span = document.createElement("span");
    span.className = className;
    if (highlight) {
      appendHighlightedText(span, part, state.searchTerm, highlight);
    } else {
      span.textContent = part;
    }
    container.appendChild(span);
  });
}

function showDefinitionTooltip(target) {
  if (!target || !elements.definitionTooltip) return;
  const term = target.dataset.term || target.textContent.trim();
  const entry = getDictionaryEntry(term);
  if (!entry) {
    const fallback = target.dataset.definition;
    if (!fallback) return;
    elements.definitionTooltipTerm.textContent = term || "Definition";
    elements.definitionTooltipMeta.classList.add("hidden");
    renderDefinitionParts(elements.definitionTooltipParts, splitDefinitionParts(fallback), {
      partClass: "definition-part"
    });
    positionDefinitionTooltip(target);
    return;
  }

  elements.definitionTooltipTerm.textContent = entry.term || term || "Definition";

  const pron = entry.pronunciation ? `(${entry.pronunciation})` : "";
  elements.definitionTooltipPron.textContent = pron;
  elements.definitionTooltipPron.classList.toggle("hidden", !entry.pronunciation);

  elements.definitionTooltipPages.textContent = entry.pages || "";
  elements.definitionTooltipPages.classList.toggle("hidden", !entry.pages);

  const hasMeta = Boolean(entry.pronunciation || entry.pages);
  elements.definitionTooltipMeta.classList.toggle("hidden", !hasMeta);

  const parts = entry.parts && entry.parts.length ? entry.parts : splitDefinitionParts(entry.definition);
  renderDefinitionParts(elements.definitionTooltipParts, parts, {
    partClass: "definition-part"
  });
  positionDefinitionTooltip(target);
}

function positionDefinitionTooltip(target) {
  const tooltip = elements.definitionTooltip;
  if (!tooltip) return;
  tooltip.classList.remove("hidden");
  tooltip.style.visibility = "hidden";
  tooltip.style.left = "0px";
  tooltip.style.top = "0px";

  const rect = target.getBoundingClientRect();
  const padding = 12;
  const tooltipRect = tooltip.getBoundingClientRect();

  let left = rect.left + window.scrollX;
  let top = rect.bottom + window.scrollY + 8;

  const maxLeft = window.scrollX + window.innerWidth - tooltipRect.width - padding;
  if (left > maxLeft) {
    left = Math.max(window.scrollX + padding, maxLeft);
  }

  if (top + tooltipRect.height > window.scrollY + window.innerHeight - padding) {
    top = rect.top + window.scrollY - tooltipRect.height - 8;
  }

  tooltip.style.left = `${Math.max(window.scrollX + padding, left)}px`;
  tooltip.style.top = `${Math.max(window.scrollX + padding, top)}px`;
  tooltip.style.visibility = "visible";
}

function hideDefinitionTooltip() {
  if (!elements.definitionTooltip) return;
  elements.definitionTooltip.classList.add("hidden");
  elements.definitionTooltip.style.left = "";
  elements.definitionTooltip.style.top = "";
  elements.definitionTooltip.style.visibility = "";
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
