/* Reader upgrade v2: source adapters, independent highlights, print-page jump.
   Existing reader code, personal notes, and study annotations remain unchanged. */
(() => {
  "use strict";
  const urls={
    bbook:"https://raw.githubusercontent.com/raemennj/bigb/main/big_book.json",
    twlvxtwlv:"https://raw.githubusercontent.com/raemennj/twlvxtwlv/main/data/twlvxtwlv.json",
    dictionary:"https://raw.githubusercontent.com/raemennj/12x12glossary/main/glossary.json",
    daily:"https://raw.githubusercontent.com/raemennj/daily/main/data/daily_reflections.json"
  };
  SOURCE_FILES.forEach(s=>{if(urls[s.id]) s.url=urls[s.id];});
  const previousNormalize=normalizeSource;
  normalizeSource=(id,data,label)=>window.AAStudyAdapters?.normalize(id,data,label)
      ||previousNormalize(id,data,label);
  const settingKey="aa-v2-highlight-layers";
  const layers={god:true,daily:true,principle:true,pronouns:false};
  try{
    const stored=JSON.parse(localStorage.getItem(settingKey)||"{}");
    Object.keys(layers).forEach(k=>{if(typeof stored[k]==="boolean")layers[k]=stored[k];});
  }catch(e){}
  const phrases=[
    "All Powerful, Guiding, Creative Intelligence","Creative Intelligence",
    "Universal Mind","Spirit of Nature","Spirit of the Universe",
    "Supreme Being","Great Reality","Infinite Power and Love",
    "God as we understood Him","God as I understood Him","God as we understand Him",
    "God as you understand God","Power greater than ourselves",
    "Power greater than myself","Power greater than yourself",
    "Power greater than himself","Power greater than themselves",
    "Power beyond ourselves","loving and All Powerful Creator",
    "Father of Light","new-found Friend","newfound Providence",
    "God of our fathers","Him who has all knowledge and power",
    "One who has all power","Prime Cause","Divine Mind","Heavenly Father",
    "our own Creator","our Creator","my Creator","our Maker",
    "new Employer","Great Spirit","His Spirit",
    "Higher Power","Creator","Maker","Providence","God","Lord"
  ];
  const regex=new RegExp("\\b(?:"+phrases.sort((a,b)=>b.length-a.length).map(escapeRegex).join("|")+")\\b","gi");
  const contextAnchor=/\b(?:God|Higher Power|Creator|Maker|Providence|Lord|Great Reality)\b/i;
  function godMarks(text){
    const marks=[];let match;
    if(layers.god){
      regex.lastIndex=0;
      while((match=regex.exec(text))){
        marks.push({start:match.index,end:match.index+match[0].length,
          className:"god-reference",priority:6,title:"God / Higher Power reference"});
      }
    }
    if(layers.pronouns){
      const pronouns=/\b(?:He|Him|His|Himself|Thou|Thee|Thy)\b/g;
      while((match=pronouns.exec(text))){
        const before=text.slice(Math.max(0,match.index-180),match.index);
        if(!contextAnchor.test(before))continue;
        marks.push({start:match.index,end:match.index+match[0].length,
          className:"god-pronoun",priority:4,title:"Possible reference — check antecedent"});
      }
    }
    return marks;
  }
  const baseHighlight=buildHighlightMatches;
  buildHighlightMatches=(text,term,options={})=>{
    const input=String(text||"");
    const others=baseHighlight(input,term,options).filter(m=>
      (layers.daily||m.className!=="crossref") &&
      (layers.principle||m.className!=="principle"));
    const candidates=others.concat(godMarks(input));
    if(!candidates.length)return [];
    const ends=[...new Set(candidates.flatMap(m=>[m.start,m.end]))].sort((a,b)=>a-b);
    const result=[];
    for(let i=1;i<ends.length;i++){
      const start=ends[i-1],end=ends[i];
      if(start===end)continue;
      const covered=candidates.filter(m=>m.start<=start&&m.end>=end);
      if(!covered.length)continue;
      covered.sort((a,b)=>(b.priority||0)-(a.priority||0));
      const cls=[...new Set(covered.flatMap(m=>m.className.split(/\s+/)))].join(" ");
      const title=[...new Set(covered.map(m=>m.title).filter(Boolean))].join(" · ");
      const dictionary=covered.find(m=>m.entry);
      const item={start,end,className:cls,title,entry:dictionary?.entry,term:dictionary?.term};
      const prev=result[result.length-1];
      if(prev&&prev.end===start&&prev.className===cls&&prev.title===title&&prev.entry===item.entry)prev.end=end;
      else result.push(item);
    }
    return result;
  };
  let pageControl;
  function indexedPage(n){return Number.isInteger(n)&&n>0;}
  function goToPrintedPage(page){
    if(!pageControl||!indexedPage(page))return;
    const source=state.sourceById.get(state.activeSourceId);
    if(!source)return;
    const pages=[];
    source.sections.forEach(sec=>(sec.pageNumbers||[]).forEach((p,i)=>{
      if(indexedPage(p))pages.push({number:p,section:sec,paragraph:i});
    }));
    const exact=pages.find(p=>p.number===page);
    const nearest=exact||pages.filter(p=>p.number<page).sort((a,b)=>b.number-a.number)[0];
    if(!nearest){
      pageControl.status.textContent="No indexed text at that page";
      return;
    }
    state.activeSectionBySource.set(source.id,nearest.section.id);
    renderActiveSource();
    pageControl.input.value=String(page);
    pageControl.status.textContent=exact?"Print page "+page:"Closest indexed page "+nearest.number;
    window.requestAnimationFrame(()=>{
      const selector='#content .print-page-marker[data-print-page="'+nearest.number+'"]';
      document.querySelector(selector)?.scrollIntoView({behavior:"smooth",block:"start"});
    });
  }
  const baseBookRender=renderBookSection;
  renderBookSection=(source,section)=>{
    baseBookRender(source,section);
    if(!pageControl)return;
    pageControl.form.hidden=!["bbook","steps","traditions"].includes(source.id);
    if(pageControl.form.hidden||!section)return;
    const paragraphs=[...elements.content.querySelectorAll(".section-block > p")];
    const pages=section.pageNumbers||[];
    let previous=null,count=0;
    paragraphs.forEach((paragraph,i)=>{
      const p=pages[i];if(p===undefined||p===null)return;
      paragraph.dataset.printPage=String(p);
      if(previous===p)return;
      previous=p;count++;
      const marker=document.createElement("div");
      marker.className="print-page-marker";
      marker.dataset.printPage=String(p);
      const label=document.createElement("span");
      label.textContent="PRINT PAGE "+p;
      marker.appendChild(label);
      paragraph.before(marker);
    });
    const first=pages.find(p=>p!==null&&p!==undefined);
    if(first!==undefined)pageControl.input.value=String(first);
    pageControl.status.textContent=count+" indexed pages in this chapter";
  };
  const baseActiveRender=renderActiveSource;
  renderActiveSource=()=>{
    baseActiveRender();
    if(pageControl&&!["bbook","steps","traditions"].includes(state.activeSourceId))
      pageControl.form.hidden=true;
  };
  function setupPages(){
    const header=document.querySelector(".reader-header");
    if(!header)return;
    const form=document.createElement("form");
    form.className="print-page-form";
    form.hidden=true;
    const label=document.createElement("label");label.htmlFor="print-page-input";label.textContent="Printed page";
    const prev=document.createElement("button");prev.type="button";prev.textContent="‹";prev.setAttribute("aria-label","Previous printed page");
    const input=document.createElement("input");input.id="print-page-input";input.type="number";input.inputMode="numeric";input.min="1";input.placeholder="62";
    const next=document.createElement("button");next.type="button";next.textContent="›";next.setAttribute("aria-label","Next printed page");
    const submit=document.createElement("button");submit.type="submit";submit.textContent="Go";
    const status=document.createElement("small");status.className="print-page-status";status.setAttribute("role","status");
    form.append(label,prev,input,next,submit,status);header.appendChild(form);
    pageControl={form,input,status};
    form.addEventListener("submit",e=>{e.preventDefault();goToPrintedPage(Number(input.value));});
    prev.addEventListener("click",()=>goToPrintedPage(Number(input.value||1)-1));
    next.addEventListener("click",()=>goToPrintedPage(Number(input.value||1)+1));
  }
  function setupHighlightControls(){
    const target=document.querySelector(".toolbar-right");if(!target)return;
    const group=document.createElement("fieldset");group.className="highlight-layers";
    const legend=document.createElement("legend");legend.textContent="Highlight layers";
    group.appendChild(legend);
    [
      ["god","God / Higher Power","coral"],
      ["daily","Daily Reflections excerpts","green"],
      ["principle","Spiritual principles","violet"],
      ["pronouns","Contextual pronouns (review)","coral"]
    ].forEach(([key,label,color])=>{
      const wrapper=document.createElement("label");wrapper.className="highlight-layer";
      wrapper.dataset.color=color;
      const checkbox=document.createElement("input");checkbox.type="checkbox";checkbox.checked=layers[key];
      const caption=document.createElement("span");caption.textContent=label;
      checkbox.addEventListener("change",()=>{
        layers[key]=checkbox.checked;
        try{localStorage.setItem(settingKey,JSON.stringify(layers));}catch(e){}
        renderActiveSource();
      });
      wrapper.append(checkbox,caption);group.appendChild(wrapper);
    });
    target.appendChild(group);
  }
  document.addEventListener("DOMContentLoaded",()=>{setupPages();setupHighlightControls();});
})();
