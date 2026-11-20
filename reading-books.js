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
  const controls = document.createElement("label");
  controls.className = "book-position";
  controls.append("Page");
  const picker = document.createElement("select");
  picker.setAttribute("aria-label", "Jump to page in " + source.title);
  controls.append(picker);
  const pages = [];
  const pageNodes = document.createDocumentFragment();
  const painted = new Set();
  function paint(index) {
    if (painted.has(index)) return;
    painted.add(index);
    const page = source.sections[index];
    const nodes = pages[index].querySelectorAll(".para");
    page.paragraphs.forEach((text, p) => appendHighlightedText(nodes[p], text, state.searchTerm, {
      includeCrossrefs: true, bookQuotes: source.quoteRanges.get(page.id + "/" + p) || []
    }));
  }
  for (const page of source.sections) {
    picker.add(new Option(page.heading, page.id));
    const block = document.createElement("section");
    block.className = "book-leaf";
    const heading = document.createElement("h3");
    heading.textContent = page.heading;
    block.append(heading);
    page.paragraphs.forEach((text, p) => {
      const para = document.createElement("p");
      para.className = "para";
      para.id = buildParagraphId(source.key, page.index, p);
      para.textContent = text;
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
  const picker = document.getElementById("bookPicker");
  if (picker.options.length === 1) {
    for (const [id, title] of readingBooks) picker.add(new Option(title, id));
    picker.onchange = () => { if (picker.value) { closeDropdown(); setActiveSource(picker.value); } };
  }
  for (const option of Array.from(picker.options).slice(1)) option.disabled = !state.sourceById.has(option.value);
  picker.value = readingBooks.has(state.activeSourceId) ? state.activeSourceId : "";
}
