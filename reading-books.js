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

function bookPageControls(source, section) {
  const controls = document.createElement("div");
  controls.className = "book-pages";
  controls.style.cssText = "display:flex;align-items:center;gap:.5rem;flex-wrap:wrap";
  function openPage(id) {
    state.activeSectionBySource.set(source.id, id);
    setActiveSource(source.id);
    document.querySelector(".book-pages select")?.focus();
  }
  const picker = document.createElement("select");
  picker.setAttribute("aria-label", "Page in " + source.title);
  picker.style.maxWidth = "100%";
  for (const page of source.sections) picker.add(new Option(page.heading, page.id));
  picker.value = section.id;
  picker.onchange = () => openPage(picker.value);
  function button(label, offset) {
    const el = document.createElement("button");
    el.type = "button";
    el.textContent = label;
    const target = source.sections[section.index + offset];
    el.disabled = !target;
    if (target) el.onclick = () => openPage(target.id);
    return el;
  }
  controls.append(picker, button("Previous page", -1), button("Next page", 1));
  return controls;
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
