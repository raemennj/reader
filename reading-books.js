const readingBooks = new Map([
  ["as-bill-sees-it", "As Bill Sees It"],
  ["came-to-believe", "Came to Believe"],
  ["comes-of-age", "Alcoholics Anonymous Comes of Age"],
  ["dr-bob", "Dr. Bob and the Good Oldtimers"]
]);
for (const [id, label] of readingBooks) SOURCE_FILES.push({ id, label, url: "books/" + id + ".json" });

function bookWords(text) {
  return Array.from(text.matchAll(/[\p{L}\p{N}]+/gu), m => ({
    word: m[0].toLowerCase(), start: m.index, end: m.index + m[0].length
  }));
}

// Match excerpt fragments across paragraph and page breaks while retaining display offsets.
function indexBookQuotations(source, daily) {
  if (source.quoteSource === daily && source.quoteRanges) return;
  source.quoteSource = daily;
  source.quoteRanges = new Map();
  if (!daily) return;
  const seeds = new Map();
  for (const section of daily.sections) {
    const { quote, date } = section.meta;
    if (!quote || !date) continue;
    for (const fragment of quote.split(/…|\.(?:\s*\.){2,}/)) {
      const words = bookWords(fragment).map(token => token.word);
      if (words.length < 4) continue;
      const width = Math.min(12, words.length);
      for (let i = 0; i <= words.length - width; i++) {
        const key = words[i];
        if (!seeds.has(key)) seeds.set(key, []);
        seeds.get(key).push({ words, i, width, date });
      }
    }
  }
  const tokens = [];
  for (const section of source.sections) {
    section.paragraphs.forEach((text, p) => {
      for (const token of bookWords(text)) tokens.push({ ...token, key: section.id + "/" + p });
    });
  }
  for (let pos = 0; pos < tokens.length; pos++) {
    for (const seed of seeds.get(tokens[pos].word) || []) {
      const { words, i, width, date } = seed;
      if (i && pos && words[i-1] === tokens[pos-1].word) continue;
      let n = 1;
      while (i+n < words.length && pos+n < tokens.length && words[i+n] === tokens[pos+n].word) n++;
      if (n < width) continue;
      for (let k = pos; k < pos+n;) {
        const first = tokens[k++];
        let last = first;
        while (k < pos+n && tokens[k].key === first.key) last = tokens[k++];
        if (!source.quoteRanges.has(first.key)) source.quoteRanges.set(first.key, []);
        source.quoteRanges.get(first.key).push({start:first.start, end:last.end, className:"crossref",
          priority:2, title:"Daily Reflection: " + date});
      }
    }
  }
}

let closeContinuousBook = () => {};
let bookAnchor = null;

function renderContinuousBook(source, selected) {
  indexBookQuotations(source, state.sourceById.get("daily"));
  elements.activeKicker.textContent = "";
  elements.activeTitle.textContent = source.title;
  const controls = document.createElement("div");
  controls.className = "book-position";
  controls.append("Page");
  const picker = document.createElement("select");
  picker.setAttribute("aria-label", "Jump to page in " + source.title);
  controls.append(picker);
  const pages = [];
  const pageNodes = document.createDocumentFragment();
  const painted = new Set();
  const printedPages = new Map();
  if (source.id === "as-bill-sees-it") {
    for (const page of source.sections) {
      const number = /^Page (\d+) · /.exec(page.heading)?.[1];
      if (number) printedPages.set(number, page);
    }
  }
  function pageLink(label, target, paragraphIndex) {
    const link = document.createElement("a");
    link.href = "#" + (paragraphIndex === undefined ? target.id : buildParagraphId(source.key, target.index, paragraphIndex));
    link.textContent = label;
    link.onclick = event => {
      event.preventDefault();
      jump(target.index);
      const leaf = pages[target.index];
      const destination = paragraphIndex === undefined ? leaf.querySelector("h3") : leaf.querySelectorAll(".para")[paragraphIndex];
      if (paragraphIndex !== undefined) {
        destination.tabIndex = -1;
        destination.scrollIntoView({block:"start", behavior:"instant"});
      }
      destination.focus({preventScroll:true});
    };
    return link;
  }
  const contents = source.sections.find(page => page.paragraphRoles?.includes("contents"));
  if (contents) controls.append(pageLink("Contents", contents));
  function paint(index) {
    if (painted.has(index)) return;
    painted.add(index);
    const page = source.sections[index];
    const nodes = pages[index].querySelectorAll(".para");
    page.paragraphs.forEach((text, p) => {
      if (nodes[p].querySelector("a")) return;
      appendHighlightedText(nodes[p], text, state.searchTerm, {
      includeCrossrefs: true, bookQuotes: source.quoteRanges.get(page.id + "/" + p) || []
      });
    });
  }
  for (const page of source.sections) {
    const subheading = page.paragraphRoles?.indexOf("subheading") ?? -1;
    const caption = subheading >= 0 ? page.heading + " · " + page.paragraphs[subheading] : page.heading;
    picker.add(new Option(caption, page.id));
    const block = document.createElement("section");
    block.className = "book-leaf";
    block.id = page.id;
    const heading = document.createElement("h3");
    heading.textContent = page.heading;
    heading.tabIndex = -1;
    block.append(heading);
    page.paragraphs.forEach((text, p) => {
      const role = page.paragraphRoles?.[p];
      const para = document.createElement(role === "subheading" ? "h4" : "p");
      para.className = "para";
      if (["source", "divider", "contents", "subheading"].includes(role)) para.classList.add("book-" + role);
      para.id = buildParagraphId(source.key, page.index, p);
      const explicit = page.contentsTargets?.[p];
      const target = role === "contents" && (explicit ? source.sectionById.get(explicit[0]) : printedPages.get(/(\d+)$/.exec(text)?.[1]));
      if (target) para.append(pageLink(text, target, explicit?.[1]));
      else para.textContent = text;
      block.append(para);
    });
    pages.push(block);
    pageNodes.append(block);
  }
  elements.content.append(controls, pageNodes);
  picker.value = selected.id;
  paint(selected.index);
  const margin = () => controls.getBoundingClientRect().height + 12;
  function jump(index) {
    paint(index);
    pages[index].scrollIntoView({block:"start", behavior:"instant"});
    state.activeSectionBySource.set(source.id, source.sections[index].id);
    picker.value = source.sections[index].id;
  }
  picker.onchange = () => jump(picker.selectedIndex);
  const restore = bookAnchor;
  bookAnchor = null;
  if (restore?.source === source && restore.page === selected.id && restore.reading) {
    window.scrollBy(0, pages[selected.index].getBoundingClientRect().top - restore.top);
  } else if (selected.index > 0) jump(selected.index);

  // Keep text and search targets in the document. Apply study markup near the viewport.
  function pageAt(y) {
    let lo = 0, hi = pages.length - 1;
    while (lo < hi) {
      const mid = Math.ceil((lo + hi) / 2);
      if (pages[mid].getBoundingClientRect().top <= y) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }
  let frame = 0;
  function sync() {
    frame = 0;
    const last = pages.length - 1;
    const end = pages[last].getBoundingClientRect();
    const current = end.top < window.innerHeight && end.bottom <= window.innerHeight ? last : pageAt(margin());
    const from = pageAt(-window.innerHeight);
    const to = pageAt(window.innerHeight * 2);
    for (let i = from; i <= to; i++) paint(i);
    picker.selectedIndex = current;
    state.activeSectionBySource.set(source.id, source.sections[current].id);
  }
  function schedule() { if (!frame) frame = requestAnimationFrame(sync); }
  window.addEventListener("scroll", schedule, {passive:true});
  window.addEventListener("resize", schedule);
  schedule();
  closeContinuousBook = () => {
    const i = pageAt(margin());
    bookAnchor = {source, page:source.sections[i].id, top:pages[i].getBoundingClientRect().top,
      reading:controls.getBoundingClientRect().top <= 1};
    window.removeEventListener("scroll", schedule);
    window.removeEventListener("resize", schedule);
    cancelAnimationFrame(frame);
    closeContinuousBook = () => {};
  };
}

function syncBookPicker() {
  const menu = document.getElementById("literatureMenu");
  if (!menu) return;
  const list = document.getElementById("literatureList");
  const trigger = menu.querySelector("summary");
  list.replaceChildren();

  function openReading(id, section, paragraph) {
    menu.open = false;
    closeDropdown();
    if (section) state.activeSectionBySource.set(id, section.id);
    setActiveSource(id);
    const target = paragraph == null ? (id === "daily" ? elements.content : elements.activeTitle) :
      document.getElementById(buildParagraphId(state.sourceById.get(id).key, section.index, paragraph));
    if (target) {
      target.tabIndex = -1;
      target.focus({preventScroll:true});
      if (section || !readingBooks.has(id)) target.scrollIntoView({block:"start"});
    }
  }

  function button(label, action) {
    const item = document.createElement("button");
    item.type = "button";
    item.textContent = label;
    item.onclick = action;
    return item;
  }

  for (const [id, title] of [["steps", "Twelve Steps"], ["traditions", "Twelve Traditions"],
      ["bbook", "Big Book"], ["daily", "Daily Reflections"], ...readingBooks]) {
    const source = state.sourceById.get(id);
    const chapters = [];
    if (source && readingBooks.has(id) && id !== "as-bill-sees-it") {
      for (const page of source.sections) {
        for (const [index, [pageId, paragraph]] of Object.entries(page.contentsTargets || {})) {
          const target = source.sectionById.get(pageId);
          if (target) chapters.push([page.paragraphs[index], target, paragraph]);
        }
      }
    } else if (source && ["steps", "traditions", "bbook"].includes(id)) {
      for (const section of source.sections) chapters.push([formatSectionLabel(section), section]);
    }
    let row;
    if (chapters.length) {
      const group = document.createElement("details");
      group.name = "literature-chapters";
      row = document.createElement("summary");
      row.textContent = title;
      group.append(row);
      for (const [label, section, paragraph] of chapters) {
        group.append(button(label, () => openReading(id, section, paragraph)));
      }
      list.append(group);
    } else {
      row = button(title, () => openReading(id));
      row.disabled = !source;
      list.append(row);
    }
    if (id === state.activeSourceId) row.setAttribute("aria-current", "true");
  }
  if (!menu.dataset.bound) {
    menu.dataset.bound = "true";
    menu.addEventListener("toggle", () => {
      if (menu.open) closeDropdown();
      else list.querySelectorAll("details[open]").forEach(group => group.open = false);
    });
    document.addEventListener("click", event => {
      if (!menu.contains(event.target)) menu.open = false;
    }, true);
    menu.addEventListener("keydown", event => {
      if (event.key === "Escape") {
        menu.open = false;
        trigger.focus();
      }
    });
  }
}
