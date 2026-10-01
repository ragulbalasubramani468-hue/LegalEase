let allTemplates={}, currentType="employment", lastId=null, lastDownloads={};

const $=id=>document.getElementById(id);
async function init(){
  allTemplates=await fetch("/api/templates").then(r=>r.json());
  renderTemplates(); renderFields();
}
function renderTemplates(){
  $("templateGrid").innerHTML="";
  for(const [key,t] of Object.entries(allTemplates)){
    const b=document.createElement("button"); b.className="template"+(key===currentType?" active":"");
    b.innerHTML=`<b>${t.name}</b><span>${t.description}</span>`;
    b.onclick=()=>{currentType=key;renderTemplates();renderFields();};
    $("templateGrid").appendChild(b);
  }
}
function renderFields(){
  const t=allTemplates[currentType]; $("fields").innerHTML="";
  for(const [key,label,placeholder] of t.fields){
    const wrap=document.createElement("div"); wrap.className="field-group";
    wrap.innerHTML=`<label>${label}</label><input data-key="${key}" placeholder="${placeholder}" value="${placeholder}">`;
    $("fields").appendChild(wrap);
  }
}
function collect(){
  const values={};
  document.querySelectorAll("#fields input").forEach(x=>values[x.dataset.key]=x.value);
  return values;
}
function esc(s){return String(s??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m]));}
function renderDoc(d){
  $("preview").innerHTML=`<h1>${esc(d.title)}</h1><div class="brandline">${esc(d.brand)}</div><div class="meta">${esc(d.date)} · ${esc(d.parties)}</div>`+
    d.sections.map(x=>`<h3>${esc(x[0])}</h3><p>${esc(x[1])}</p>`).join("");
  $("terms").innerHTML=d.terms.map(x=>`<div class="term"><span>${esc(x[0])}</span><b>${esc(x[1])}</b></div>`).join("");
}
$("generateBtn").onclick=async()=>{
  const btn=$("generateBtn");btn.disabled=true;btn.innerHTML="Generating…";
  const req={doc_type:currentType,values:collect(),brand_name:$("brandName").value,logo_url:$("logoUrl").value,font:$("font").value,title:$("title").value,extra_terms:$("extraTerms").value};
  try{
    const res=await fetch("/api/generate",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(req)}).then(r=>r.json());
    lastId=res.id;lastDownloads=res.download;renderDoc(res.document);
    ["pdfBtn","docxBtn","txtBtn"].forEach(id=>$(id).disabled=false);
  }catch(e){alert("Could not generate the document.");}
  btn.disabled=false;btn.innerHTML='Generate document <span>→</span>';
};
function dl(type){if(lastDownloads[type]) location.href=lastDownloads[type];}
$("pdfBtn").onclick=()=>dl("pdf");$("docxBtn").onclick=()=>dl("docx");$("txtBtn").onclick=()=>dl("txt");
$("copyBtn").onclick=async()=>{await navigator.clipboard.writeText($("preview").innerText);$("copyBtn").textContent="Copied";setTimeout(()=>$("copyBtn").textContent="Copy",1200)};
$("printBtn").onclick=()=>window.print();
$("newBtn").onclick=()=>{lastId=null;$("preview").innerHTML='<div class="empty"><div class="empty-icon">✦</div><h3>Your document will appear here</h3><p>Select a template and generate a draft.</p></div>';["pdfBtn","docxBtn","txtBtn"].forEach(id=>$(id).disabled=true);window.scrollTo({top:0,behavior:"smooth"});};
init();
