/* AA Study Library — optional, independent study annotations.
   Annotations are editorial study interpretations, not A.A. publication headings. */
(function () {
  "use strict";
  var lens = { data: null, chapter: "all", type: "turn", query: "", enabled: true, open: false };
  var ui = {};
  var STORAGE_KEY = "aa-study-lens-enabled-v1";
  var TYPE_ALL = "all";
  function create(tag, className, textValue) {
    var node = document.createElement(tag);
    if (className) node.className = className;
    if (textValue !== undefined) node.textContent = textValue;
    return node;
  }
  function normalize(value) {
    return String(value || "").toLowerCase().normalize("NFKD")
      .replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]+/g, " ").trim();
  }
  function chapterForSection(section) {
    if (!lens.data || !section) return null;
    var heading = normalize(section.heading);
    return lens.data.chapters.find(function (chapter) {
      var title = normalize(chapter.title);
      return heading === title || heading.includes(title) || title.includes(heading);
    }) || null;
  }
  function currentChapter() {
    try {
      if (typeof state === "undefined" || state.activeSourceId !== "bbook") return null;
      var src = state.sourceById.get("bbook");
      var sid = state.activeSectionBySource.get("bbook");
      return chapterForSection(src && src.sectionById.get(sid));
    } catch (e) { return null; }
  }
  function getType(type) {
    return lens.data.types.find(function (entry) { return entry.id === type; });
  }
  function show() {
    lens.open = true;
    ui.modal.classList.add("lens-open");
    ui.modal.setAttribute("aria-hidden", "false");
    document.body.classList.add("lens-modal-open");
    redraw();
    ui.close.focus();
  }
  function hide() {
    lens.open = false;
    ui.modal.classList.remove("lens-open");
    ui.modal.setAttribute("aria-hidden", "true");
    document.body.classList.remove("lens-modal-open");
    ui.launch.focus();
  }
  function displayMessage(message) {
    ui.message.textContent = message;
  }
  function goToEntry(chapter, entry) {
    if (typeof state === "undefined" || typeof renderActiveSource !== "function") {
      displayMessage("Reader navigation is not available in this version.");
      return;
    }
    var src = state.sourceById.get("bbook");
    if (!src) {
      displayMessage("Load bbook.json (via Choose JSON files) to jump into the Big Book.");
      return;
    }
    var sec = src.sections.find(function (candidate) {
      return chapterForSection(candidate) && chapterForSection(candidate).id === chapter.id;
    });
    if (!sec) {
      displayMessage("That chapter is not present in the uploaded Big Book file.");
      return;
    }
    state.activeSourceId = "bbook";
    state.activeSectionBySource.set("bbook", sec.id);
    renderActiveSource();
    hide();
    window.requestAnimationFrame(function () {
      var anchor = normalize(entry.anchor);
      var paragraphs = Array.from(document.querySelectorAll("#content p.para"));
      var candidates = anchor.length >= 9
        ? paragraphs.filter(function (p) { return normalize(p.textContent).includes(anchor); })
        : [];
      var destination = candidates.length === 1 ? candidates[0] : document.querySelector("#content");
      if (destination) destination.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  }
  function makeFilter(type, value, label) {
    var btn = create("button", "lens-filter", label);
    btn.type = "button";
    if ((type === "type" && lens.type === value) || (type === "chapter" && lens.chapter === value)) {
      btn.classList.add("selected");
      btn.setAttribute("aria-pressed", "true");
    } else {
      btn.setAttribute("aria-pressed", "false");
    }
    if (type === "type" && value !== TYPE_ALL) {
      btn.style.setProperty("--lens-color", getType(value).color);
    }
    btn.addEventListener("click", function () {
      lens[type] = value;
      redraw();
    });
    return btn;
  }
  function redraw() {
    if (!lens.data) return;
    ui.typeFilters.replaceChildren();
    ui.typeFilters.appendChild(makeFilter("type", TYPE_ALL, "All"));
    lens.data.types.forEach(function (type) {
      ui.typeFilters.appendChild(makeFilter("type", type.id, type.label));
    });
    var chapters = lens.data.chapters.filter(function (ch) {
      return lens.chapter === TYPE_ALL || lens.chapter === ch.id;
    });
    var needle = normalize(lens.query);
    ui.list.replaceChildren();
    var count = 0;
    chapters.forEach(function (ch) {
      var entries = ch.entries.filter(function (entry) {
        return (lens.type === TYPE_ALL || entry.type === lens.type)
          && (!needle || normalize(ch.title + " " + entry.label + " " + entry.note + " " + entry.page).includes(needle));
      });
      if (!entries.length) return;
      var group = create("section", "lens-group");
      var groupTitle = create("div", "lens-group-title", ch.title + "  ·  pp. " + ch.start + "–" + ch.end);
      group.appendChild(groupTitle);
      entries.forEach(function (entry) {
        count++;
        var typ = getType(entry.type);
        var card = create("article", "lens-entry");
        card.style.setProperty("--lens-color", typ.color);
        var top = create("div", "lens-entry-top");
        var page = create("span", "lens-page", "p. " + entry.page);
        var kind = create("span", "lens-kind", typ.label);
        top.append(page, kind);
        var title = create("strong", "lens-entry-title", entry.label);
        var note = create("p", "lens-entry-note", entry.note);
        var go = create("button", "lens-go", "Open chapter ↗");
        go.type = "button";
        go.addEventListener("click", function () { goToEntry(ch, entry); });
        card.append(top, title, note, go);
        group.appendChild(card);
      });
      ui.list.appendChild(group);
    });
    ui.total.textContent = count + " matching study markers";
    if (!count) ui.list.appendChild(create("p", "lens-empty", "No markers match these filters."));
    ui.toggle.checked = lens.enabled;
    ui.chapterSelect.value = lens.chapter;
  }
  function refreshReaderMarkers() {
    if (!lens.data) return;
    var header = document.querySelector(".reader-header");
    if (!header) return;
    var previous = header.querySelector(".lens-reader-strip");
    if (previous) previous.remove();
    var paragraphs = Array.from(document.querySelectorAll("#content p.para, #content p.section-subtitle"));
    paragraphs.forEach(function (p) {
      p.classList.remove("lens-passage");
      p.style.removeProperty("--lens-color");
      p.removeAttribute("data-lens-title");
      p.querySelectorAll(".lens-paragraph-badge").forEach(function (badge) { badge.remove(); });
    });
    if (!lens.enabled) return;
    var chapter = currentChapter();
    if (!chapter) return;
    var strip = create("div", "lens-reader-strip");
    var title = create("span", "lens-strip-caption", "Study Lens · " + chapter.entries.length + " markers");
    strip.appendChild(title);
    ["problem", "turn", "solution", "action"].forEach(function (type) {
      var entries = chapter.entries.filter(function (entry) { return entry.type === type; });
      if (!entries.length) return;
      var chip = create("button", "lens-strip-chip", getType(type).label + " " + entries.length);
      chip.type = "button";
      chip.style.setProperty("--lens-color", getType(type).color);
      chip.addEventListener("click", function () {
        lens.chapter = chapter.id;
        lens.type = type;
        show();
      });
      strip.appendChild(chip);
    });
    var all = create("button", "lens-strip-all", "See chapter map →");
    all.type = "button";
    all.addEventListener("click", function () {
      lens.chapter = chapter.id;
      lens.type = TYPE_ALL;
      show();
    });
    strip.appendChild(all);
    header.appendChild(strip);
    chapter.entries.forEach(function (entry) {
      var anchor = normalize(entry.anchor);
      if (anchor.length < 9) return;
      var matches = paragraphs.filter(function (p) {
        return normalize(p.textContent).includes(anchor);
      });
      // Only decorate unambiguous paragraph matches, never guess at an exact line.
      if (matches.length !== 1) return;
      var p = matches[0];
      if (p.classList.contains("lens-passage")) return;
      var typ = getType(entry.type);
      p.classList.add("lens-passage");
      p.style.setProperty("--lens-color", typ.color);
      p.dataset.lensTitle = "Study suggestion · " + typ.label + " · p. " + entry.page + ": " + entry.label;
      var badge = create("button", "lens-paragraph-badge", typ.label + " · p. " + entry.page);
      badge.type = "button";
      badge.title = entry.label + " — " + entry.note + " (page reference is editorial)";
      badge.addEventListener("click", function () {
        lens.chapter = chapter.id;
        lens.type = entry.type;
        show();
      });
      p.appendChild(badge);
    });
  }
  function installStyles() {
    var style = create("style");
    style.textContent = [
      ".lens-modal-open{overflow:hidden}.lens-launch{white-space:nowrap}",
      ".lens-modal{position:fixed;inset:0;z-index:9990;display:none;align-items:center;justify-content:center;padding:16px}",
      ".lens-modal.lens-open{display:flex}.lens-backdrop{position:absolute;inset:0;background:rgba(15,20,23,.62)}",
      ".lens-dialog{position:relative;z-index:1;background:#faf7f1;color:#252621;border:1px solid #d4cfc5;border-radius:17px;width:min(900px,100%);max-height:90vh;box-shadow:0 22px 90px #17161466;display:flex;flex-direction:column;overflow:hidden}",
      ".lens-dialog-head{display:flex;justify-content:space-between;align-items:start;gap:15px;padding:16px 20px 12px}",
      ".lens-dialog-head h2{margin:0;font-family:Georgia,serif;font-size:1.8rem}.lens-subhead{font-size:.87rem;color:#635e55;margin:4px 0 0}",
      ".lens-close{border:1px solid #c6c0b6;border-radius:12px;background:white;padding:7px 14px;cursor:pointer}",
      ".lens-controls{padding:4px 20px 14px;border-bottom:1px solid #e3ddd3;display:grid;gap:11px}",
      ".lens-control-row{display:flex;align-items:center;gap:9px;flex-wrap:wrap}",
      ".lens-control-row select,.lens-control-row input[type=search]{min-height:36px;max-width:100%;border:1px solid #c9c2b5;border-radius:9px;padding:6px 9px;background:white;color:#222}",
      ".lens-control-row input[type=search]{flex:1;min-width:150px}",
      ".lens-control-row label{font-size:.94rem}",
      ".lens-type-filters{display:flex;flex-wrap:wrap;gap:6px}",
      ".lens-filter{background:#fff;border:1px solid #d5cec4;border-radius:20px;padding:5px 10px;cursor:pointer;font-family:inherit;color:#333}",
      ".lens-filter.selected{background:#212d2a;color:white;border-color:#212d2a;box-shadow:inset 0 0 0 2px var(--lens-color,transparent)}",
      ".lens-total{font-size:.82rem;color:#666}",".lens-list{overflow:auto;padding:12px 20px 22px}",
      ".lens-group{margin:8px 0 18px}.lens-group-title{font-size:1.02rem;font-weight:700;color:#4d4a43;padding:8px 0}",
      ".lens-entry{margin:7px 0;padding:11px 13px;background:white;border:1px solid #e4e0d9;border-left:6px solid var(--lens-color);border-radius:8px;display:grid;gap:5px}",
      ".lens-entry-top{display:flex;gap:8px;align-items:center}.lens-page{font-weight:800;font-size:.9rem}.lens-kind{font-size:.78rem;color:#666}",
      ".lens-entry-title{font-size:1.04rem}.lens-entry-note{font-size:.92rem;margin:0;line-height:1.4;color:#4d4b46}",
      ".lens-go{justify-self:start;padding:2px 0;border:0;background:none;color:#17634e;text-decoration:underline;cursor:pointer}",
      ".lens-footnote{margin:0;padding:8px 20px 12px;color:#6b655e;border-top:1px solid #ddd6cc;font-size:.8rem;line-height:1.3}",
      ".lens-message{font-size:.81rem;color:#a34c3f;min-height:1em}",
      ".lens-reader-strip{display:flex;gap:6px;align-items:center;flex-wrap:wrap;margin:7px 0 12px;padding:9px 0;border-top:1px solid #ded9cc;border-bottom:1px solid #ded9cc}",
      ".lens-strip-caption{font-size:.86rem;font-weight:700;margin-right:5px}.lens-strip-chip{border:1px solid #d6d3cb;border-radius:14px;background:var(--lens-color);color:#2d2925;font-size:.78rem;padding:4px 9px;cursor:pointer}",
      ".lens-strip-all{margin-left:auto;border:0;background:transparent;text-decoration:underline;color:#286651;cursor:pointer;font-size:.83rem}",
      "#content p.lens-passage{border-left:4px solid var(--lens-color)!important;padding-left:12px!important;background:color-mix(in srgb,var(--lens-color) 16%,transparent);border-radius:3px}",
      ".lens-paragraph-badge{display:block;margin:6px 0 0 auto;border:1px solid #9d9993;border-radius:10px;background:#fff9;padding:3px 8px;font:12px system-ui;color:#302d29;cursor:pointer}",
      "@media(max-width:600px){.lens-modal{padding:6px}.lens-dialog{max-height:97dvh}.lens-dialog-head h2{font-size:1.48rem}.lens-controls{padding:6px 12px}.lens-dialog-head{padding:12px}.lens-list{padding:8px 12px}.lens-footnote{padding:7px 12px}.lens-strip-all{margin-left:0}}"
    ].join("\n");
    document.head.appendChild(style);
  }
  function init() {
    var actions = document.querySelector(".data-actions") || document.querySelector(".toolbar-right");
    if (!actions) return;
    try { lens.enabled = localStorage.getItem(STORAGE_KEY) !== "false"; } catch (e) {}
    installStyles();
    ui.launch = create("button", "ghost lens-launch", "Study Lens");
    ui.launch.type = "button";
    actions.insertBefore(ui.launch, actions.firstChild);
    ui.modal = create("div", "lens-modal");
    ui.modal.setAttribute("aria-hidden", "true");
    var backdrop = create("div", "lens-backdrop");
    backdrop.addEventListener("click", hide);
    var dialog = create("div", "lens-dialog");
    dialog.setAttribute("role", "dialog");
    dialog.setAttribute("aria-modal", "true");
    dialog.setAttribute("aria-label", "Study Lens");
    var head = create("div", "lens-dialog-head");
    var headingBox = create("div");
    headingBox.append(create("h2", "", "Problem → Solution → Action"),
      create("p", "lens-subhead", "Print-page study map • Big Book, Fourth Edition • independently prepared"));
    ui.close = create("button", "lens-close", "Close");
    ui.close.type = "button";
    ui.close.addEventListener("click", hide);
    head.append(headingBox, ui.close);
    var controls = create("div", "lens-controls");
    var firstRow = create("div", "lens-control-row");
    ui.chapterSelect = create("select");
    var allOpt = create("option", "", "All chapters");
    allOpt.value = "all";
    ui.chapterSelect.appendChild(allOpt);
    ui.chapterSelect.addEventListener("change", function () {
      lens.chapter = ui.chapterSelect.value; redraw();
    });
    ui.search = create("input");
    ui.search.type = "search";
    ui.search.placeholder = "Search markers and pages";
    ui.search.setAttribute("aria-label", "Search study markers");
    ui.search.addEventListener("input", function () { lens.query = ui.search.value; redraw(); });
    var toggleLabel = create("label");
    ui.toggle = create("input");
    ui.toggle.type = "checkbox";
    ui.toggle.addEventListener("change", function () {
      lens.enabled = ui.toggle.checked;
      try { localStorage.setItem(STORAGE_KEY, String(lens.enabled)); } catch (e) {}
      refreshReaderMarkers();
    });
    toggleLabel.append(ui.toggle, document.createTextNode(" Mark reader passages"));
    firstRow.append(ui.chapterSelect, ui.search, toggleLabel);
    ui.typeFilters = create("div", "lens-type-filters");
    ui.total = create("div", "lens-total");
    ui.message = create("div", "lens-message");
    ui.message.setAttribute("role", "status");
    controls.append(firstRow, ui.typeFilters, ui.total, ui.message);
    ui.list = create("div", "lens-list");
    var footnote = create("p", "lens-footnote",
      "Editorial study prompts, not official A.A. classifications. Page numbers indicate print pages to inspect. Colored paragraph cues are shown only when a text anchor matches uniquely; they do not verify exact page placement.");
    dialog.append(head, controls, ui.list, footnote);
    ui.modal.append(backdrop, dialog);
    document.body.appendChild(ui.modal);
    ui.launch.addEventListener("click", show);
    document.addEventListener("keydown", function (event) {
      if (event.key === "Escape" && lens.open) hide();
    });
    var reader = document.getElementById("content");
    if (reader) {
      new MutationObserver(function () {
        window.requestAnimationFrame(refreshReaderMarkers);
      }).observe(reader, { childList: true });
    }
    fetch("study-map.json", { cache: "no-cache" })
      .then(function (res) { if (!res.ok) throw new Error("Study map not found"); return res.json(); })
      .then(function (data) {
        if (!data.chapters || !data.types) throw new Error("Invalid study map");
        lens.data = data;
        data.chapters.forEach(function (ch) {
          var option = create("option", "", ch.title + " (pp. " + ch.start + "–" + ch.end + ")");
          option.value = ch.id;
          ui.chapterSelect.appendChild(option);
        });
        ui.launch.textContent = "Study Lens · " + data.chapters.reduce(function (n, ch) { return n + ch.entries.length; }, 0);
        redraw();
        refreshReaderMarkers();
      }).catch(function (e) {
        ui.list.textContent = "Could not load study-map.json: " + e.message;
        ui.launch.textContent = "Study Lens (unavailable)";
      });
  }
  document.addEventListener("DOMContentLoaded", init);
})();
