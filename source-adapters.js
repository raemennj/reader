/* AA Study Library v2 — source adapters.
 * Converts the owner's existing public JSON documents into the reader's section model.
 * Big Book page numbers come from printed_page, NOT PDF index. */
(function () {
  "use strict";
  function cleanLine(line) {
    return (line.spans || []).map(function (s) { return s.text || ""; }).join("").normalize("NFKC");
  }
  function blockText(block) {
    if (block.text) return String(block.text);
    var lines = (block.lines || []).map(cleanLine).filter(Boolean);
    return lines.join(" ").replace(/\s+/g, " ").trim();
  }
  function indexSections(sections) {
    sections.forEach(function (s, idx) {
      s.index = idx;
    });
    return new Map(sections.map(function (s) { return [s.id, s]; }));
  }
  var bbChapters = [
    [1, "Bill's Story"], [17, "There Is a Solution"], [30, "More About Alcoholism"],
    [44, "We Agnostics"], [58, "How It Works"], [72, "Into Action"],
    [89, "Working With Others"], [104, "To Wives"], [122, "The Family Afterward"],
    [136, "To Employers"], [151, "A Vision for You"], [165, "Personal Stories and Appendices"]
  ];
  function makeBigBook(id, data) {
    if (!Array.isArray(data.pages)) return null;
    var sections = [{id:"bbook-front", heading:"Front Matter & The Doctor's Opinion", paragraphs:[], pageNumbers:[], meta:{type:"front"}}]
      .concat(bbChapters.map(function (chapter, i) {
        return {id:"bbook-" + chapter[0],heading:chapter[1], paragraphs:[],pageNumbers:[],meta:{type:"chapter",number:i+1,startPage:chapter[0]}};
      }));
    data.pages.forEach(function (page) {
      var pageNum = /^\d+$/.test(String(page.printed_page || "")) ? Number(page.printed_page) : null;
      var which = 0;
      if (pageNum !== null) {
        bbChapters.forEach(function (entry, i) { if (pageNum >= entry[0]) which = i + 1; });
      }
      var sec = sections[which];
      (page.blocks || []).forEach(function (block) {
        if (["page_number", "running_head"].includes(block.type)) return;
        var txt = blockText(block);
        if (!txt) return;
        sec.paragraphs.push(txt);
        sec.pageNumbers.push(pageNum !== null ? pageNum : (page.printed_page || null));
      });
    });
    var used = sections.filter(function (s) { return s.paragraphs.length > 0; });
    return {id:"bbook",title:"Big Book",kind:"book",metadata:"Alcoholics Anonymous · Fourth Edition",author:"Alcoholics Anonymous World Services",sections:used,sectionById:indexSections(used)};
  }
  function makeTwelve(_id, data) {
    if (!Array.isArray(data.sections) || !data.sections.some(function (s) { return s.type === "heading"; })) return null;
    var steps = [], traditions = [];
    var current = null, category = "steps", i=0;
    data.sections.forEach(function (b) {
      if (b.type === "heading") {
        var name = String(b.text || "").trim();
        if (/^THE TWELVE TRADITIONS$/i.test(name)) { category="traditions";current=null;return; }
        if (/^THE TWELVE STEPS$/i.test(name)) { category="steps";current=null;return; }
        if (/^(Step|Tradition)\s+(One|Two|Three|Four|Five|Six|Seven|Eight|Nine|Ten|Eleven|Twelve)$/i.test(name) || /^Foreword$/i.test(name) || /^Traditions.?.?Long Form$/i.test(name)) {
          category = /^Tradition/i.test(name) ? "traditions" : category;
          current={id:"aa12-"+(++i),heading:name,paragraphs:[],pageNumbers:[],meta:{type:category==="steps"?"step":"tradition",startPage:b.pageNumber||null}};
          (category==="steps"?steps:traditions).push(current);
        }
        return;
      }
      if (!current) return;
      var text = String(b.text || "");
      if (b.type === "list") text += " " + (b.items||[]).join("; ");
      if (text) {current.paragraphs.push(text);current.pageNumbers.push(b.pageNumber||null);}
    });
    function make(id,title,sections) {
      return {id:id,title:title,kind:"book",metadata:"Twelve Steps and Twelve Traditions",author:"Alcoholics Anonymous World Services",sections:sections,sectionById:indexSections(sections)};
    }
    return [make("steps","Twelve Steps",steps),make("traditions","Twelve Traditions",traditions)];
  }
  function makeDictionary(id, data) {
    if (!data.entries || typeof data.entries !== "object") return null;
    var sections = Object.entries(data.entries).map(function (pair,i) {
      var name = pair[1].word||pair[1].name||pair[0];
      var def = pair[1].definition||pair[1].meaning||"";
      return {id:"glossary-"+i,heading:name,paragraphs:[def],pageNumbers:[],
        meta:{type:"dictionary",pronunciation:pair[1].pronunciation||"",pages:(pair[1].pages||[]).join(", "),definitionParts:[def]}};
    }).sort(function(a,b){return a.heading.localeCompare(b.heading);});
    return {id:"dictionary",title:"Dictionary",kind:"book",metadata:"12 & 12 Glossary",author:"Independent companion",sections:sections,sectionById:indexSections(sections)};
  }
  window.AAStudyAdapters = {
    normalize:function (id,data) {
      if (id==="bbook" && data && Array.isArray(data.pages)) return makeBigBook(id,data);
      if (id==="twlvxtwlv" && data && Array.isArray(data.sections) && data.metadata) return makeTwelve(id,data);
      if (id==="dictionary" && data && data.entries) return makeDictionary(id,data);
      return null;
    }
  };
})();
