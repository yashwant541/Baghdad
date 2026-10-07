/* Bahrain-Iraq Recon Studio — frontend.
   No inline event handlers anywhere (CSP-safe): every interaction is wired
   through addEventListener + event delegation, and every network call goes
   through api() which never lets a non-JSON / error response fail silently. */

let state={session:null,pivot:[],pivotCols:[],tbTree:[],tbRaw:[],tbItems:[],tbSel:new Set(),acctView:'desc',localCcy:'IQD',fileTypes:{},depth:null,defaultRun:null,workLang:null,outstanding:null,uiLang:'en',
  tbGroupItems:[],tbSelGroups:new Set(),tbLevel:'account',
  subFilesMeta:[],selection:{},subPreview:[],subSel:new Set(),subExtra:[],
  suggestions:[],suggestionDecisions:{},manualMatches:[],results:[],lineage:[],
  mappingCoverage:[],mappingCoverageSummary:null,coverageFilter:'all',
  amountSearchResults:[],amountSearchAdded:new Set()};

const $=id=>document.getElementById(id);

function backendUrl(p){
  try{ return typeof getWebAppBackendUrl==='function' ? getWebAppBackendUrl(p) : p; }
  catch(e){ console.error('getWebAppBackendUrl failed, falling back to raw path',e); return p; }
}

/* ---------------- safe networking ---------------- */

async function api(path,options){
  let res;
  try{
    res=await fetch(backendUrl(path),options);
  }catch(networkErr){
    throw new Error(`Could not reach the backend (${path}). ${networkErr.message||networkErr}`);
  }
  let text='';
  try{ text=await res.text(); }catch(e){ /* ignore */ }
  let data=null;
  if(text){
    try{ data=JSON.parse(text); }
    catch(parseErr){
      throw new Error(`Backend returned an unexpected (non-JSON) response for ${path} — status ${res.status}. `+
        `This usually means the server raised an error. Check the Dataiku backend log.`);
    }
  }
  if(!res.ok){
    throw new Error((data&&data.error)||`Request to ${path} failed with status ${res.status}`);
  }
  if(data&&data.ok===false){
    throw new Error(data.error||`Request to ${path} was rejected by the server`);
  }
  return data||{};
}

/* ---------------- small utilities ---------------- */

function setStatus(s,ok){$('runStatus').textContent=s;const d=$('statusDot');if(d)d.style.background=ok===false?'var(--danger)':'var(--success)'}
function toast(s){let t=$('toast');t.textContent=s;t.style.display='block';clearTimeout(toast._h);toast._h=setTimeout(()=>t.style.display='none',3800)}
function go(id){
  const page=$(id); if(!page) return;
  document.querySelectorAll('.page,.step').forEach(x=>x.classList.remove('active'));
  page.classList.add('active');
  const nav=document.querySelector(`.step[data-page="${id}"]`); if(nav) nav.classList.add('active');
}
function fmt(v){return typeof v==='number'?v.toLocaleString(undefined,{maximumFractionDigits:2}):(v??'')}
function esc(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function table(id,rows,cols){let t=$(id);if(!rows||!rows.length){t.innerHTML='<tr><td>No records yet</td></tr>';return}
  t.innerHTML='<thead><tr>'+cols.map(c=>`<th>${esc(c.label||c.key)}</th>`).join('')+'</tr></thead><tbody>'+
  rows.map(r=>'<tr>'+cols.map(c=>`<td>${c.render?c.render(r[c.key],r):esc(fmt(r[c.key]))}</td>`).join('')+'</tr>').join('')+'</tbody>'}

/* every button/action runs through here: disables the control while busy,
   and turns thrown errors into a visible toast instead of a silent no-op */
async function guarded(el,fn){
  const busyBtn = el && el.tagName==='BUTTON' ? el : null;
  // (the button's HTML is never swapped out, so the interface translator keeps its original English text)
  if(busyBtn){ busyBtn.disabled=true; busyBtn.classList.add('isBusy'); }
  try{
    await fn();
  }catch(err){
    console.error(err);
    toast((err&&err.message)||String(err));
    setStatus('Error — see message',false);
  }finally{
    if(busyBtn){ busyBtn.disabled=false; busyBtn.classList.remove('isBusy'); }
  }
}

/* ---------------- init & global error safety net ---------------- */

async function init(){
  try{
    const j=await api('/api/session',{method:'POST'});
    state.session=j.session_id;
    setStatus('Session ready');
  }catch(err){
    console.error(err);
    setStatus('Could not start a session',false);
    toast(err.message||String(err));
  }
}
window.addEventListener('error',e=>{console.error('Unhandled error',e.error||e.message)});
window.addEventListener('unhandledrejection',e=>{console.error('Unhandled promise rejection',e.reason)});

/* ---------------- Trial Balance ---------------- */

async function uploadTB(){
  const f=$('tbFile').files[0]; if(!f){toast('Choose a Trial Balance workbook first');return}
  const fd=new FormData(); fd.append('session_id',state.session); fd.append('file',f);
  setStatus('Processing Trial Balance');
  const j=await api('/api/tb',{method:'POST',body:fd});
  state.pivot=j.preview; state.pivotCols=j.columns; state.tbTree=j.tree;
  state.localCcy=j.local_currency||'IQD';
  if(j.engine_has_frx===false) toast('The bahrain_iraq_algorithm library in your project is out of date (no FRX* support) — copy the latest engine.py into it and restart the backend.');
  state.tbRaw=flattenTbItems(state.tbTree); state.tbItems=aggregateAccounts(state.tbRaw,state.acctView,null,true); state.tbSel=new Set();
  state.tbGroupItems=flattenTbGroups(state.tbTree); state.tbSelGroups=new Set();
  $('tbMessage').innerHTML=`<div class="notice good">Detected sheet <b>${esc(j.meta.sheet)}</b>, header row <b>${j.meta.header_row}</b>, ${j.meta.rows} data rows.</div>`;
  $('kpis').innerHTML=Object.entries(j.kpis).map(([k,v])=>`<div class="kpi"><small>${esc(k.replace('_',' '))}</small><b>${esc(fmt(v))}</b></div>`).join('');
  renderPivot(); renderTbTree(); renderTbPicker(); renderDepthFiles(); setStatus('TB pivot ready'); go('pivot');
}
function renderPivot(){
  const q=($('pivotSearch')?.value||'').toLowerCase();
  const rows=(state.pivot||[]).filter(r=>JSON.stringify(r).toLowerCase().includes(q));
  table('pivotTable',rows,(state.pivotCols||[]).map(c=>({key:c,label:c})));
}
// FRX is the synthetic "Forex" aggregate: every currency other than local (IQD)
function ccyLabel(c){ return c==='FRX'?'FRX*':c; }
function currencyChips(totals){
  if(!totals) return '';
  return Object.entries(totals).filter(([k])=>k!=='TOTAL').map(([k,v])=>`<span class="chip${k==='FRX'?' chipClub':''}" ${k==='FRX'?'title="Forex: sub-total of every currency other than the local currency"':''}>${esc(ccyLabel(k))} ${esc(fmt(v))}</span>`).join('')+
    `<span class="chip chipTotal">TOTAL ${esc(fmt(totals.TOTAL))}</span>`;
}
function renderTbTree(){
  const html=(state.tbTree||[]).map(g=>`
    <details class="treeNode groupNode" open>
      <summary><span class="nodeName">${esc(g.name)}</span>${currencyChips(g.currency_totals)}</summary>
      <div class="treeChildren">${(g.children||[]).map(c=>`
        <details class="treeNode subNode">
          <summary><span class="nodeName">${esc(c.name)}</span>${currencyChips(c.currency_totals)}<span class="muted">${(c.accounts||[]).length} account rows</span></summary>
          <div class="accountList">${(c.accounts||[]).map(a=>`<div class="acctRow"><span>${esc(a.account)} — ${esc(a.account_desc)}</span><span class="muted">${esc(a.currency)}</span><b>${esc(fmt(a.amount))}</b></div>`).join('')||'<div class="muted" style="padding:8px">No account rows</div>'}</div>
        </details>`).join('')}
        ${(g.accounts||[]).length?`<div class="accountList">${g.accounts.map(a=>`<div class="acctRow"><span>${esc(a.account)} — ${esc(a.account_desc)}</span><span class="muted">${esc(a.currency)}</span><b>${esc(fmt(a.amount))}</b></div>`).join('')}</div>`:''}
      </div>
    </details>`).join('');
  $('tbTreePreview').innerHTML=html||'<div class="muted">No structure detected yet.</div>';
}

/* ---------------- Submissions ---------------- */

function initSelection(){
  state.selection={};
  state.subFilesMeta.forEach(f=>{ state.selection[f.file]={};
    (f.sheets||[]).forEach(s=>{ state.selection[f.file][s.sheet]={checked: s.header_row!=null && s.score>=3, header_row: s.header_row}; }); });
}
function revealSheetPicker(){
  renderSheetPicker();
  $('extractActions').style.display='flex';
  $('subDownloadRow').style.display='none';
}
async function uploadSubs(){
  const fs=[...$('subFiles').files]; if(!fs.length){toast('Choose submission files first');return}
  const fd=new FormData(); fd.append('session_id',state.session); fs.slice(0,8).forEach(f=>fd.append('files',f));
  setStatus('Inspecting submission workbooks');
  const j=await api('/api/submissions/upload',{method:'POST',body:fd});
  state.subFilesMeta=j.files||[]; state.workLang=null;
  initSelection();
  state.fileTypes={}; state.subFilesMeta.forEach(f=>{ state.fileTypes[f.file]=f.guessed_type||'any'; });
  state.depth=null; renderDepthFiles(); renderDepthResults();
  $('subDownloadRow').style.display='none';
  if(j.has_arabic){
    // the submissions are (partly) Arabic: ask whether to translate or work in Arabic before showing any sheet
    $('sheetPicker').innerHTML=''; $('extractActions').style.display='none';
    $('subMessage').innerHTML='';
    showLangChoice();
    setStatus('Choose a working language');
    return;
  }
  $('langChoice').style.display='none';
  revealSheetPicker();
  $('subMessage').innerHTML=`<div class="notice good">Inspected ${state.subFilesMeta.length} workbook(s). Choose the sheets to read, then extract.</div>`;
  setStatus('Choose submission sheets');
}
function showLangChoice(){
  const rows=state.subFilesMeta.filter(f=>f.arabic_cells||((f.arabic_sheets||[]).length));
  $('langChoiceText').innerHTML=rows.map(f=>`<div><b>${esc(f.file)}</b> — ${esc(f.arabic_cells)} Arabic cell(s) in ${esc((f.arabic_sheets||[]).length)} sheet(s)${(f.arabic_sample||[]).length?` <span class="muted">(${esc(f.arabic_sample.slice(0,2).join(' · '))})</span>`:''}</div>`).join('');
  document.querySelectorAll('input[name="workLang"]').forEach(r=>{ r.checked=false; });
  $('dictBox').style.display='none'; $('dictName').textContent=''; $('dictFile').value='';
  $('langChoice').style.display='block';
}
async function proceedLang(){
  const sel=document.querySelector('input[name="workLang"]:checked');
  if(!sel){toast('Choose Arabic or English first');return}
  if(sel.value==='ar'){
    await api('/api/submissions/language',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({session_id:state.session,mode:'ar'})});
    state.workLang='ar'; $('langChoice').style.display='none'; revealSheetPicker();
    $('subMessage').innerHTML='<div class="notice good">Working in Arabic — the submissions are used exactly as uploaded. Choose the sheets to read, then extract.</div>';
    setStatus('Choose submission sheets');
    return;
  }
  const f=$('dictFile').files[0]; if(!f){toast('Choose your dictionary workbook first');return}
  const fd=new FormData(); fd.append('session_id',state.session); fd.append('dictionary',f);
  setStatus('Translating the submissions to English');
  const j=await api('/api/submissions/translate',{method:'POST',body:fd});
  state.subFilesMeta=j.files||[]; initSelection(); state.workLang='en';
  state.subFilesMeta.forEach(x=>{ if(!state.fileTypes[x.file]) state.fileTypes[x.file]=x.guessed_type||'any'; });
  state.depth=null; renderDepthFiles(); renderDepthResults();
  $('langChoice').style.display='none'; revealSheetPicker();
  const cells=j.stats.reduce((a,x)=>a+x.translated_cells,0), sheets=j.stats.reduce((a,x)=>a+x.translated_sheets,0);
  const dl=j.stats.filter(x=>!x.skipped).map(x=>`<button type="button" class="secondary" data-action="download-translated" data-file="${esc(x.file)}">Download ${esc(x.file)} (translated)</button>`).join(' ');
  $('subMessage').innerHTML=`<div class="notice good">Translated ${cells} cell(s) and ${sheets} sheet name(s) using ${j.entries} dictionary entries. ${j.missing_count} Arabic term(s) were not in the dictionary — the English workbook is used from here on.
    <div style="margin-top:10px">${j.missing_count?'<button type="button" class="secondary" data-action="download-missing">Download missing terms</button> ':''}${dl}</div></div>`;
  setStatus('Choose submission sheets');
}

/* ---------------- Off-balance: the Outstanding Report and its two pivots ---------------- */

async function uploadOutstanding(){
  const f=$('obFile').files[0]; if(!f){toast('Choose the Outstanding Report workbook first');return}
  const fd=new FormData(); fd.append('session_id',state.session); fd.append('file',f);
  setStatus('Building the Outstanding Report pivots');
  const j=await api('/api/outstanding/upload',{method:'POST',body:fd});
  state.outstanding=j; state.depth=null; renderOutstanding(); renderDepthFiles(); renderDepthResults();
  setStatus('Outstanding Report ready');
}
async function changeObCategory(value){
  const j=await api('/api/outstanding/category',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({session_id:state.session,category:value})});
  state.outstanding=j; state.depth=null; renderOutstanding(); renderDepthFiles(); renderDepthResults();
}
function obCell(v){
  if(typeof v!=='number') return esc(v);
  if(Math.abs(v)<0.005) return '-';
  return v<0?`<span class="negNum">(${esc(fmt(Math.abs(v)))})</span>`:esc(fmt(v));
}
function renderObTable(id,flat){
  const t=$(id); if(!t) return;
  const nl=flat.label_cols.length;
  const head='<thead><tr>'+flat.label_cols.concat(flat.columns).map(c=>`<th>${esc(c)}</th>`).join('')+'</tr></thead>';
  const rank={'Category':0,'LC/GTEE':1,'Currency':2,'Grand Total':0};
  const body=flat.rows.map(r=>{
    const bold=r.level==='Category'||r.level==='Grand Total'||(r.level==='LC/GTEE'&&nl===3);
    const labs=r.labels.map((l,i)=>nl===3?(i===rank[r.level]?l:''):l);
    return `<tr class="${bold?'obBold':''} ${r.level==='Grand Total'?'obGrand':''}">`+labs.map(l=>`<td data-tr="1">${esc(l)}</td>`).join('')+
      r.values.concat([r.total]).map(v=>`<td class="num">${obCell(v)}</td>`).join('')+'</tr>';
  }).join('');
  t.innerHTML=head+'<tbody>'+body+'</tbody>';
}
function renderOutstanding(){
  const j=state.outstanding; if(!j) return;
  const val={}; (j.validation||[]).forEach(([k,v])=>{ val[k]=v; });
  $('obMessage').innerHTML=`<div class="notice good">Detected sheet <b>${esc(j.sheet)}</b>, header row <b>${esc(j.header_row)}</b> in <b>${esc(j.source)}</b>.</div>`+
    (j.warnings||[]).map(w=>`<div class="notice">${esc(w)}</div>`).join('');
  $('obKpis').innerHTML=[['Rows read',val['Source rows read']],['Detail rows',val['Cleaned detail rows']],
    ['Invalid / blank Equ-IQD',(val['Rows with invalid / non-numeric Equ-IQD (excluded)']||0)+(val['Rows with blank Equ-IQD (excluded)']||0)],
    ['Source total',val['Source Equ-IQD total (valid rows)']],['Difference',val['Reconciliation difference (source - Pivot 1)']]]
    .map(([k,v])=>`<div class="kpi"><small>${esc(k)}</small><b>${esc(fmt(v))}</b></div>`).join('');
  const sel=$('obCategory');
  sel.innerHTML=['<option value="">All</option>'].concat((j.categories||[]).map(c=>`<option value="${esc(c)}" ${c===j.selected_category?'selected':''}>${esc(c)}</option>`)).join('');
  renderObTable('obPivot1',j.pivot1); renderObTable('obPivot2',j.pivot2);
  $('obValidation').innerHTML='<thead><tr><th>Check</th><th>Value</th></tr></thead><tbody>'+(j.validation||[]).map(([k,v])=>`<tr><td data-tr="1">${esc(k)}</td><td>${esc(typeof v==='number'?fmt(v):v)}</td></tr>`).join('')+'</tbody>';
  $('obBody').style.display='block'; $('btnOutstandingDl').style.display='inline-flex';
}

/* ---------------- Interface language (English / Arabic) ---------------- */

function setLang(lang){
  state.uiLang=lang==='ar'?'ar':'en';
  try{ localStorage.setItem('recon_ui_lang',state.uiLang); }catch(e){}
  if(typeof applyLanguage==='function') applyLanguage();
}
function renderSheetPicker(){
  $('sheetPicker').innerHTML=state.subFilesMeta.map(f=>`
    <div class="fileCard">
      <div class="fileCardHead"><b>${esc(f.file)}</b><span class="muted">${(f.sheets||[]).length} sheet(s)</span></div>
      ${(f.sheets||[]).map(s=>{
        const sel=state.selection[f.file][s.sheet];
        const ccy=Object.entries(s.currency_columns||{}).map(([col,info])=>`<span class="chip">${esc(col)}: ${esc(info.currency)}${info.scale>1?` ×${info.scale}`:''}</span>`).join('')||'<span class="muted">No currency columns detected — will use raw numeric cells</span>';
        return `<label class="sheetRow">
          <input type="checkbox" data-action="toggle-sheet" data-file="${esc(f.file)}" data-sheet="${esc(s.sheet)}" ${sel.checked?'checked':''}>
          <div class="sheetInfo">
            <div class="sheetTitle">${esc(s.sheet)} <span class="muted">${s.rows}×${s.columns}</span></div>
            <div class="sheetChips">${ccy}</div>
          </div>
          <div class="sheetHeaderRow"><small>Header row</small><input type="number" min="1" data-action="set-header-row" data-file="${esc(f.file)}" data-sheet="${esc(s.sheet)}" value="${sel.header_row??''}"></div>
        </label>`}).join('')}
    </div>`).join('');
}
async function extractSubs(){
  const selection={};
  for(const file in state.selection){
    const sheets=[];
    for(const sheet in state.selection[file]){
      const s=state.selection[file][sheet];
      if(s.checked) sheets.push(s.header_row?{sheet,header_row:s.header_row}:{sheet});
    }
    if(sheets.length) selection[file]=sheets;
  }
  if(!Object.keys(selection).length){toast('Select at least one sheet');return}
  setStatus('Extracting submission sheets');
  const j=await api('/api/submissions/extract',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({session_id:state.session,selection})});
  state.subPreview=j.preview||[]; state.subSel=new Set();
  state.suggestions=j.suggestions||[];
  // when a group/currency has more than one candidate suggestion (e.g. the
  // blanket "sum everything" version and a narrower auto-clubbed subset
  // alternative), only ever default-check the single best-fitting one —
  // approving both would submit two conflicting rules for the same target.
  const bestForKey={};
  state.suggestions.forEach((s,i)=>{
    const key=s.bs_mapping+'|'+s.currency;
    if(bestForKey[key]==null||Math.abs(s.difference)<Math.abs(state.suggestions[bestForKey[key]].difference)) bestForKey[key]=i;
  });
  state.suggestionDecisions={};
  state.suggestions.forEach((s,i)=>{
    const key=s.bs_mapping+'|'+s.currency;
    state.suggestionDecisions[i]=(bestForKey[key]===i)&&s.confidence>=0.6;
  });
  // the bundled default mapping (see load_default_mapping) is resolved
  // server-side against this TB + these submissions automatically — pull
  // whatever it found straight into Manual matches, same as a manually
  // uploaded mapping, just without the upload step. Still fully reviewable
  // (and removable) before you ever run the reconciliation.
  const defaultMatches=j.default_mapping_matches||[];
  state.manualMatches=state.manualMatches.filter(m=>m.source!=='DEFAULT');
  defaultMatches.forEach(r=>state.manualMatches.push({...r.resolved,warnings:r.warnings||[]}));
  state.defaultRun=j.default_results||null; renderDefault();
  state.mappingCoverage=j.mapping_coverage||[];
  state.mappingCoverageSummary=j.mapping_coverage_summary||null;
  const defaultNote=defaultMatches.length
    ? ` Included <b>${defaultMatches.length}</b> match${defaultMatches.length!==1?'es':''} from the default mapping automatically (review under Manual matches — see Mapping Coverage for the full picture: ${state.mappingCoverageSummary?state.mappingCoverageSummary.fulfilled+' of '+state.mappingCoverageSummary.total+' rules fulfilled':''}).`
    : '';
  $('subExtractMessage').innerHTML=`<div class="notice good">Extracted <b>${j.count}</b> currency-tagged lines. Generated <b>${state.suggestions.length}</b> candidate matches.${defaultNote}</div>`;
  $('subDownloadRow').style.display=state.subPreview.length?'flex':'none';
  table('subTable',state.subPreview,[
    {key:'submission_file',label:'File'},{key:'sheet',label:'Sheet'},{key:'source_cell',label:'Cell'},
    {key:'hierarchy_path',label:'Section'},{key:'line_description',label:'Line description'},
    {key:'currency',label:'Currency'},{key:'is_total',label:'Total row?',render:v=>v?'<span class="pill MATCH">TOTAL</span>':''},
    {key:'normalized_amount',label:'Amount'}]);
  renderSuggestions(); renderSubPicker(); renderTbPicker(); renderManualMatches(); renderCoverage();
  setStatus('Submissions extracted'); go('defaultmap');
}
// a plain `window.location = url` download works in a normal browser tab,
// but Dataiku serves this webapp inside a sandboxed iframe, where top-level
// navigation like that is commonly blocked outright (silently, or with a
// console warning nobody sees) — "downloading nothing ever happens" is
// exactly what that looks like. fetch + blob + a programmatic <a download>
// click works inside a sandboxed iframe as long as downloads are allowed at
// all, and — unlike a raw navigation — lets an error response show up as a
// proper toast instead of navigating the whole app to a page of raw JSON.
async function downloadFile(path,fallbackName,opts){
  const res=await fetch(backendUrl(path),opts);
  if(!res.ok){
    let msg=`Download failed (status ${res.status})`;
    try{ const data=await res.json(); if(data&&data.error) msg=data.error; }catch(e){}
    throw new Error(msg);
  }
  const blob=await res.blob();
  let filename=fallbackName;
  const cd=res.headers.get('Content-Disposition');
  if(cd){ const m=/filename\*?=(?:UTF-8'')?"?([^";]+)"?/.exec(cd); if(m) filename=decodeURIComponent(m[1]); }
  const a=document.createElement('a');
  a.href=URL.createObjectURL(blob); a.download=filename;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(a.href);
}
async function downloadSubmissions(){
  if(!state.session){toast('No active session yet');return}
  await downloadFile('/api/submissions/export?session_id='+encodeURIComponent(state.session),'Submissions_Simplified.xlsx');
}

/* ---------------- Default Mapping: every rule, per file, with its variance ---------------- */

const DM_STATUS={MATCH:'Match',MATCH_WITHIN_TOLERANCE:'Match (within tolerance)',REVIEW_REQUIRED:'Variance - review',NOT_FOUND:'Not found'};
function dmPct(v){ return v==null?'':(v.toLocaleString(undefined,{maximumFractionDigits:4})+'%'); }
function dmFileSheets(file){ const f=state.subFilesMeta.find(x=>x.file===file); return new Set((f?f.sheets:[]).map(s=>s.sheet)); }
function dmRulesFor(file,all){
  const sheets=dmFileSheets(file);
  return all.filter(r=>(r.files||[]).includes(file)||(!(r.files||[]).length&&(r.target_sheets||[]).some(s=>sheets.has(s))));
}
function dmSummary(rs){
  const s={rules:rs.length,MATCH:0,MATCH_WITHIN_TOLERANCE:0,REVIEW_REQUIRED:0,NOT_FOUND:0,abs:0};
  rs.forEach(r=>{ s[r.status]++; if(r.status==='REVIEW_REQUIRED'||r.status==='MATCH_WITHIN_TOLERANCE') s.abs+=Math.abs(r.variance); });
  return s;
}
function dmChips(s){
  return [`${s.rules} rules`,`${s.MATCH+s.MATCH_WITHIN_TOLERANCE} matched`,`${s.REVIEW_REQUIRED} with a variance to review`,`${s.NOT_FOUND} not found`]
    .map(c=>`<span class="chip">${esc(c)}</span>`).join('');
}
function renderDefault(){
  const j=state.defaultRun, box=$('dmFiles'); if(!box) return;
  const ready=!!(j&&j.ok);
  $('dmPrereq').style.display=ready?'none':'block';
  $('dmBody').style.display=ready?'block':'none';
  if(j&&!j.ok){ $('dmPrereq').style.display='block'; $('dmPrereq').textContent=j.error||'The default mapping could not be run.'; return; }
  if(!ready){ box.innerHTML=''; return; }
  const all=j.results, s=j.summary;
  $('dmKpis').innerHTML=[['Rules applied',s.rules,''],['Match',s.MATCH,''],['Match (within tolerance)',s.MATCH_WITHIN_TOLERANCE,''],
      ['Variance - review',s.REVIEW_REQUIRED,s.REVIEW_REQUIRED?' bad':''],['Not found',s.NOT_FOUND,'']]
    .map(([l,v,c])=>`<div class="kpi dmKpi${c}"><small>${esc(l)}</small><b>${esc(v)}</b></div>`).join('');
  const nf=all.filter(r=>r.status==='NOT_FOUND').length;
  const note=$('dmNote');
  note.style.display=s.REVIEW_REQUIRED?'block':'none';
  note.textContent=`Total absolute variance on rules that matched with a difference: ${fmt(s.abs_variance)}. Variance = Trial Balance amount - submission amount.`;
  const q=($('dmFilter').value||'').toLowerCase(), st=$('dmStatus').value;
  const keep=r=>(!st||r.status===st)&&(!q||(r.label+' '+r.rule_text+' '+r.where+' '+(r.files||[]).join(' ')).toLowerCase().includes(q));
  const rowsHtml=rs=>rs.filter(keep).map(r=>`<tr class="${r.status==='REVIEW_REQUIRED'?'dmReview':''}">
      <td>${r.color?`<span class="swatch" style="background:#${esc(r.color)}"></span>`:''}</td><td>${esc(r.n)}</td>
      <td data-tr="1"><b>${esc(r.label)}</b> <span class="muted">${esc(ccyLabel(r.currency))}</span></td><td class="how">${esc(r.rule_text)}</td>
      <td class="num">${esc(fmt(r.tb_amount))}</td><td class="num">${esc(fmt(r.sub_amount))}</td>
      <td class="num ${r.status==='REVIEW_REQUIRED'?'bad':''}">${esc(fmt(r.variance))}</td><td class="num">${esc(dmPct(r.variance_pct))}</td>
      <td><span class="pill ${r.status}">${esc(DM_STATUS[r.status])}</span></td>
      <td>${esc(r.where||'(not found)')}${(r.warnings||[]).length?`<div class="muted">${esc(r.warnings.slice(0,2).join(' | '))}</div>`:''}</td></tr>`).join('');
  const head='<thead><tr><th></th><th>#</th><th>Rule</th><th>How the rule works</th><th>TB amount</th><th>Submission amount</th><th>Variance</th><th>Variance %</th><th>Status</th><th>Where in the file</th></tr></thead>';
  const empty='<div class="muted" style="padding:8px 2px">No rules for this filter.</div>';
  let out=state.subFilesMeta.map(f=>{
    const rs=dmRulesFor(f.file,all); if(!rs.length) return '';
    const sm=dmSummary(rs), body=rowsHtml(rs);
    return `<div class="depthCard"><div class="depthCardHead"><div><b>${esc(f.file)}</b> <span class="chip">${esc((DEPTH_TYPES.find(t=>t[0]===state.fileTypes[f.file])||[0,''])[1])}</span></div>
      <button type="button" class="primary" data-action="dm-download" data-file="${esc(f.file)}">Download annotated workbook</button></div>
      <div>${dmChips(sm)}${sm.REVIEW_REQUIRED||sm.MATCH_WITHIN_TOLERANCE?`<span class="chip">${esc('Total absolute variance '+fmt(sm.abs))}</span>`:''}</div>
      ${body?`<div class="depthScroll"><table class="dmTable">${head}<tbody>${body}</tbody></table></div>`:empty}</div>`;
  }).join('');
  const shown=new Set(); state.subFilesMeta.forEach(f=>dmRulesFor(f.file,all).forEach(r=>shown.add(r.n)));
  const orphan=all.filter(r=>!shown.has(r.n));
  if(orphan.length){
    const body=rowsHtml(orphan);
    out+=`<div class="depthCard"><div class="depthCardHead"><div><b>Rules that apply to none of the uploaded files</b></div></div>
      ${body?`<div class="depthScroll"><table class="dmTable">${head}<tbody>${body}</tbody></table></div>`:empty}</div>`;
  }
  box.innerHTML=out||'<div class="muted">No rules to show.</div>';
  $('dmNotes').innerHTML=(j.notes||[]).map(n=>`<div class="notice">${esc(n)}</div>`).join('');
  const gs=j.groups||[], open_=gs.filter(g=>!g.covered&&Math.abs(g.total)>1e-6);
  $('dmGroups').innerHTML=gs.length?`<details class="depthCard" ${open_.length?'open':''}><summary><b>Trial Balance groups and the rules that read them</b> <span class="chip">${esc(gs.length-open_.length+' of '+gs.length+' groups covered')}</span>${open_.length?`<span class="chip">${esc(open_.length+' group(s) with no rule')}</span>`:''}</summary>
    <div class="depthScroll"><table class="dmTable" style="min-width:600px"><thead><tr><th>BS mapping group</th><th>IQD</th><th>Foreign currencies</th><th>Total</th><th>Rules that read it</th></tr></thead><tbody>${gs.map(g=>`<tr class="${g.covered?'':'dmReview'}"><td><b>${esc(g.group)}</b></td><td class="num">${esc(fmt(g.iqd))}</td><td class="num">${esc(fmt(g.frx))}</td><td class="num">${esc(fmt(g.total))}</td><td>${g.covered?esc(g.rules.map(n=>'#'+n).join(', ')):'<span class="pill NOT_FOUND">NO RULE</span>'}</td></tr>`).join('')}</tbody></table></div></details>`:'';
  $('btnDmAll').style.display=j.files&&j.files.some(f=>f.summary.rules>0)?'inline-flex':'none';
}
async function rerunDefault(){
  const num=(id,def)=>{const v=parseFloat($(id).value); return isNaN(v)?def:v};
  setStatus('Re-running the default mapping');
  const j=await api('/api/default/run',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({
    session_id:state.session,tolerance_abs:num('dmTolAbs',1),tolerance_pct:num('dmTolPct',0.01)/100})});
  state.defaultRun=j; renderDefault();
  setStatus(`Default mapping: ${j.summary.REVIEW_REQUIRED} variance(s) to review`);
}
function downloadDefaultPdf(){
  const manual=state.manualMatches.filter(m=>(m.source||'MANUAL')!=='DEFAULT').map(m=>({label:m.label,bs_mapping:m.bs_mapping,currency:m.currency,
    source:m.source||'MANUAL',tb_components:(m.tb_components||[]).map(c=>({amount:c.amount,sign:c.sign})),
    components:(m.components||[]).map(c=>({amount:c.amount,sign:c.sign,multiplier:c.multiplier}))}));
  const approved=(state.suggestions||[]).map((s,i)=>!!state.suggestionDecisions[i]);
  return downloadFile('/api/default/pdf','Default_Mapping_Results.pdf',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({session_id:state.session,manual_matches:manual,approved})});
}

/* ---------------- Depth Search: TB pivot values -> every non-zero cell, every sheet ---------------- */

const DEPTH_TYPES=[['assets','Assets (033)'],['liabilities','Liabilities (034)'],
  ['offbalance','Off-balance (064)'],['any','Other — search everything']];

function renderDepthFiles(){
  const box=$('depthFiles'); if(!box) return;
  const ready=state.tbRaw.length>0 && state.subFilesMeta.length>0;
  const pre=$('depthPrereq'); if(pre) pre.style.display=ready?'none':'block';
  const ex=$('depthExtra');
  if(ex){
    const o=state.outstanding;
    ex.style.display=o?'block':'none';
    const hasOb=Object.values(state.fileTypes).some(t=>t==='offbalance');
    ex.className='notice'+(hasOb?' good':'');
    if(o) ex.innerHTML=hasOb
      ? `Outstanding Report loaded (${esc(o.sheet)}): it is searched in the Off-balance (064) file only - its Pivot 1 (Bucket) on Maturity sheets, and Pivot 2 (CATEGORY: ${esc(o.selected_category||'All')}) on every sheet.`
      : `Outstanding Report loaded (${esc(o.sheet)}), but no file is typed Off-balance (064) - set the type of the Off-balance submission below, otherwise its pivots are not searched.`;
  }
  if(!state.subFilesMeta.length){ box.innerHTML='<div class="muted" style="padding:6px 2px">No submission files yet — upload them in the Submissions step.</div>'; return; }
  box.innerHTML=state.subFilesMeta.map(f=>{
    const cur=state.fileTypes[f.file]||'any';
    return `<div class="depthFileRow"><b>${esc(f.file)}</b><span class="muted">${(f.sheets||[]).length} sheet(s)</span>
      <select data-action="set-file-type" data-file="${esc(f.file)}">${DEPTH_TYPES.map(([v,l])=>`<option value="${v}" ${v===cur?'selected':''}>${esc(l)}</option>`).join('')}</select></div>`;
  }).join('');
}
async function runDepth(){
  if(!state.tbRaw.length){toast('Upload the Trial Balance first');return}
  if(!state.subFilesMeta.length){toast('Upload the submission files first');return}
  const num=(id,def)=>{const v=parseFloat($(id).value); return isNaN(v)?def:v};
  $('depthMessage').innerHTML='';
  setStatus('Searching every sheet for Trial Balance values');
  const j=await api('/api/depth-search/run',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({
    session_id:state.session,file_types:state.fileTypes,
    params:{min_value:num('depthMin',1000),min_scaled:num('depthMinScaled',100000),tol_abs:num('depthTol',1),
            allow_scale:$('depthScale').checked,allow_sign:$('depthSign').checked,include_accounts:$('depthAccounts').checked,
            maturity_keywords:$('depthMat').value}})});
  state.depth=j; renderDepthResults();
  $('depthMessage').innerHTML=(j.notes||[]).map(n=>`<div class="notice">${esc(n)}</div>`).join('');
  const found=j.files.reduce((a,f)=>a+((f.stats||{}).targets_matched||0),0);
  setStatus(`Depth search done — ${found} TB value(s) located`);
}
function depthHow(h){
  const p=[];
  if(h.scale!==1) p.push(scaleLabel(h.scale)+(h.scale_declared?' (sheet is in thousands)':''));
  if(h.opposite_sign) p.push('opposite sign');
  return p.join(', ')||'exact';
}
function renderDepthResults(){
  const box=$('depthResults'); if(!box) return;
  const j=state.depth, filt=$('depthFilter');
  if(!j){ box.innerHTML=''; if(filt) filt.style.display='none'; $('btnDepthAll').style.display='none'; return; }
  if(filt) filt.style.display='block';
  const q=(filt?.value||'').toLowerCase();
  $('btnDepthAll').style.display=j.files.some(f=>!f.skipped)?'inline-flex':'none';
  box.innerHTML=j.files.map(f=>{
    const st=f.stats||{};
    const head=`<div class="depthCardHead"><div><b>${esc(f.file)}</b> <span class="chip">${esc(f.type_label)}</span></div>
      ${f.skipped?'':`<button type="button" class="primary" data-action="download-depth" data-file="${esc(f.file)}">Download annotated workbook</button>`}</div>`;
    if(f.skipped) return `<div class="depthCard">${head}<div class="muted">${esc((f.warnings||[]).join(' '))}</div></div>`;
    const chips=[`${st.sheets_scanned} sheets scanned`,`${fmt(st.numeric_cells)} non-zero cells`,`${st.targets_matched} of ${st.targets_searched} TB values found`]
      .concat(f.type==='offbalance'?[`Searched: ${st.from_tb_ob||0} OB- Trial Balance values + ${st.from_outstanding||0} Outstanding Report values`]:[])
      .concat(st.currency_conflicts?[`${st.currency_conflicts} currency conflict(s) not highlighted`]:[]).map(c=>`<span class="chip">${esc(c)}</span>`).join('');
    const warn=(f.warnings||[]).map(w=>`<div class="notice" style="margin:8px 0 0">${esc(w)}</div>`).join('');
    const hits=(f.hits||[]).filter(h=>{
      if(!q) return true; const t=j.targets[h.target_id]||{};
      return (t.label+' '+t.currency+' '+h.sheet+' '+h.cell+' '+h.row_label+' '+h.column_header).toLowerCase().includes(q); });
    const shown=hits.slice(0,400);
    const rows=shown.map(h=>{ const t=j.targets[h.target_id]||{};
      return `<tr class="${h.currency_check==='CONFLICT'?'conflict':''}">
        <td>${h.color?`<span class="swatch" style="background:#${esc(h.color)}"></span>`:''}</td>
        <td data-tr="1"><b>${esc(t.label)}</b> <span class="muted">${esc(ccyLabel(t.currency))}</span>${t.sheet&&t.sheet!=='TB Pivot'&&t.sheet!=='TB Accounts'?` <span class="muted">· ${esc(t.sheet)}</span>`:''}</td><td class="num">${esc(fmt(t.amount))}</td>
        <td>${esc(h.sheet)}</td><td><b>${esc(h.cell)}</b></td>
        <td class="num">${esc(h.value.toLocaleString(undefined,{maximumFractionDigits:6}))}</td><td data-tr="1">${esc(depthHow(h))}</td>
        <td>${esc(h.row_label)}</td><td>${esc(h.column_header)}</td>
        <td data-tr="1" title="${h.currency_check==='CONFLICT'?'This column is a different currency from the TB value':''}">${esc(h.currency_check)}</td></tr>`; }).join('');
    const table=hits.length?`<div class="depthScroll"><table class="depthTable"><thead><tr><th></th><th>TB item</th><th>TB amount</th><th>Sheet</th><th>Cell</th><th>Sheet value</th><th>How matched</th><th>Row label</th><th>Column</th><th>Currency</th></tr></thead><tbody>${rows}</tbody></table></div>`
      +(hits.length>shown.length?`<div class="muted" style="padding:6px 2px">Showing ${shown.length} of ${hits.length} — download the workbook for the full report.</div>`:'')
      :'<div class="muted" style="padding:8px 2px">No matching cells'+(q?' for this filter':'')+'.</div>';
    const un=(f.unmatched||[]).slice(0,500).map(u=>`<tr><td>${esc(u.level)}</td><td><b>${esc(u.label)}</b></td><td>${esc(ccyLabel(u.currency))}</td><td class="num">${esc(fmt(u.amount))}</td></tr>`).join('');
    return `<div class="depthCard">${head}<div>${chips}</div>${warn}${table}
      ${(f.unmatched||[]).length?`<details><summary>${f.unmatched.length} TB value(s) not found in this file</summary><div class="depthScroll"><table class="depthTable"><thead><tr><th>Level</th><th>TB item</th><th>Currency</th><th>TB amount</th></tr></thead><tbody>${un}</tbody></table></div></details>`:''}</div>`;
  }).join('');
}

/* ---------------- Mapping Studio: auto-suggestions ---------------- */

function renderSuggestionSummary(){
  const accepted=state.suggestions.filter((_,i)=>state.suggestionDecisions[i]);
  const groups=new Set(accepted.map(s=>s.group));
  $('suggestionSummary').innerHTML=[
    ['Suggestions',state.suggestions.length],['Approved',accepted.length],
    ['TB groups covered',groups.size],['Avg confidence',accepted.length?Math.round(100*accepted.reduce((a,s)=>a+s.confidence,0)/accepted.length)+'%':'—']
  ].map(([k,v])=>`<div class="kpi"><small>${esc(k)}</small><b>${esc(v)}</b></div>`).join('');
}
function scaleLabel(scale){
  if(!scale||scale===1) return null;
  return scale>1 ? `×${fmt(scale)}` : `÷${fmt(1/scale)}`;
}
function renderSuggestions(){
  $('suggestions').innerHTML=state.suggestions.map((s,i)=>{
    const status = Math.abs(s.difference)<1e-6?'MATCH':(Math.abs(s.difference)<=1 || Math.abs(s.difference)/Math.max(Math.abs(s.tb_amount),1)<=0.0001?'MATCH_WITHIN_TOLERANCE':'REVIEW_REQUIRED');
    const isClub = s.rule_type==='AUTO_CLUB_SUBSET'||s.rule_type==='AUTO_CLUB_SUBSET_SCALED';
    const isScaled = s.rule_type==='AUTO_SCALE_ADJUST'||s.rule_type==='AUTO_CLUB_SUBSET_SCALED';
    const sLabel = scaleLabel(s.scale_applied);
    return `<div class="suggestionCard ${state.suggestionDecisions[i]?'accepted':''} ${isClub||isScaled?'clubCard':''}">
      <div class="suggestionHead">
        <label class="acceptToggle"><input type="checkbox" data-action="toggle-suggestion" data-index="${i}" ${state.suggestionDecisions[i]?'checked':''}><span>Approve</span></label>
        <div class="suggestionTitle"><b>${esc(s.bs_mapping)}</b><span class="chip">${esc(s.currency)}</span>
          ${isClub?'<span class="chip chipClub">Auto-clubbed subset</span>':''}
          ${isScaled?`<span class="chip chipClub">Scale-adjusted ${esc(sLabel)}</span>`:''}
          <span class="pill ${status}">${status}</span></div>
        <span class="confidence">confidence ${Math.round(s.confidence*100)}%</span>
      </div>
      <div class="suggestionBody">
        <div class="amountCols">
          <div><small>TB amount</small><b>${esc(fmt(s.tb_amount))}</b></div>
          <div><small>Submission amount</small><b>${esc(fmt(s.suggested_submission_amount))}</b></div>
          <div><small>Difference</small><b class="${status==='REVIEW_REQUIRED'?'bad':'good'}">${esc(fmt(s.difference))}</b></div>
          <div><small>Sign applied</small><b>${s.sign_applied<0?'Flipped (×-1)':'As reported'}</b></div>
        </div>
        ${isClub?`<div class="matchWarnings" style="margin-top:10px">This is a narrower alternative to the blanket "sum every line" suggestion above for the same group — it excludes ${s.match_basis.excluded_lines} line(s) that didn't fit${isScaled?`, and reports the submission side ${esc(sLabel)} relative to what the sheet shows`:''}. Approve at most one of the two for this group/currency.</div>`
          :(isScaled?`<div class="matchWarnings" style="margin-top:10px">The submission side didn't reconcile at face value, but does after applying ${esc(sLabel)} — the sheet's header didn't state a scale, so this was detected purely from the numbers. Double-check before approving.</div>`:'')}
        <details class="evidence"><summary>Match basis &amp; evidence (${s.components.length} line${s.components.length!==1?'s':''})</summary>
          <div class="matchBasis muted">Section match: <b>${esc(s.match_basis.section_match||'—')}</b> · Sub-group match: <b>${esc(s.match_basis.subgroup_match||'—')}</b>${s.match_basis.multiple_hits?` · <b>${s.match_basis.multiple_hits} sheets/files had this section</b> — ${s.match_basis.picked_by_amount?'picked the one whose amount agrees with the TB':'none agreed with the TB, took the closest'}: ${esc(s.match_basis.picked_source||'')}`:''} · ${s.match_basis.used_total_rows?'Used pre‑computed Total rows':(isClub?`Auto-clubbed a subset of lines (excluded ${s.match_basis.excluded_lines})`:'Summed individual lines')}${sLabel?` · Scale detected: ${esc(sLabel)}`:''}</div>
          <table class="miniTable"><thead><tr><th>File</th><th>Sheet</th><th>Cell</th><th>Description</th><th>Sheet amount</th>${sLabel?'<th>Adjusted amount</th>':''}</tr></thead>
          <tbody>${s.components.map(c=>`<tr><td>${esc(c.submission_file)}</td><td>${esc(c.sheet)}</td><td>${esc(c.source_cell||c.row_number)}</td><td>${esc(c.line_description)}</td><td>${esc(fmt(c.amount))}</td>${sLabel?`<td>${esc(fmt(c.amount*(c.multiplier||1)))}</td>`:''}</tr>`).join('')}</tbody></table>
        </details>
      </div>
    </div>`}).join('') || '<div class="notice">No candidate matches were generated. Build a manual match below.</div>';
  renderSuggestionSummary();
}

/* ---------------- Mapping Studio: manual match builder ---------------- */

function flattenTbItems(tree){
  const out=[];
  function walk(node){
    (node.accounts||[]).forEach(a=>out.push({account:a.account,account_desc:a.account_desc,bs_mapping:a.bs_mapping,currency:a.currency,amount:a.amount}));
    (node.children||[]).forEach(walk);
  }
  (tree||[]).forEach(walk);
  return out;
}
// The same account often appears several times in a TB: one description under
// several account numbers, or one number carrying several descriptions. This
// folds raw account rows into one row per account DESCRIPTION ('desc') or per
// account NUMBER ('no'), summing amounts, for display and picking only — the
// raw rows are kept in `members` and are what the reconciliation really uses.
function normDesc(s){ return String(s==null?'':s).trim().toLowerCase().replace(/\s+/g,' '); }
function rawAcctKey(a){ return [a.account,a.account_desc,a.currency,a.bs_mapping].join('|'); }
// withFrx (picker only): also emit a synthetic "FRX*" row per account = the
// sub-total of all its non-local-currency lines. Picking it saves the mapping as
// "all foreign currencies", so a currency that is new next period is still included.
function aggregateAccounts(rows,mode,amt,withFrx){
  amt=amt||(r=>r.amount);
  const m=new Map();
  const add=(k,r,ccy,forex)=>{
    let g=m.get(k);
    if(!g){ g={account:r.account,account_desc:r.account_desc,bs_mapping:r.bs_mapping,currency:ccy,forex:!!forex,amount:0,members:[],accounts:[],descs:[],ccys:[]}; m.set(k,g); }
    g.amount+=amt(r); g.members.push(r);
    if(!g.accounts.includes(r.account)) g.accounts.push(r.account);
    if(!g.descs.includes(r.account_desc)) g.descs.push(r.account_desc);
    if(!g.ccys.includes(r.currency)) g.ccys.push(r.currency);
  };
  (rows||[]).forEach(r=>{
    const id=mode==='no'?String(r.account):(normDesc(r.account_desc)||('#'+r.account));
    add(id+'|'+r.currency+'|'+r.bs_mapping,r,r.currency,false);
    if(withFrx && r.currency!==state.localCcy) add(id+'|FRX|'+r.bs_mapping,r,'FRX',true);
  });
  return [...m.values()];
}
function aggTitle(g,mode){
  if(mode==='no'){
    const extra=g.descs.length>1?` (+${g.descs.length-1} other name${g.descs.length>2?'s':''})`:'';
    return `${g.account} — ${g.descs[0]||''}${extra}`;
  }
  return g.descs[0]||String(g.account);
}
function aggAccountsNote(g,mode){
  if(mode==='no') return '';
  const shown=g.accounts.slice(0,4).join(', ');
  return 'acct '+shown+(g.accounts.length>4?` +${g.accounts.length-4} more`:'');
}
function collectTbAccounts(node){
  let out=(node.accounts||[]).slice();
  (node.children||[]).forEach(c=>{ out=out.concat(collectTbAccounts(c)); });
  return out;
}
// one selectable row per (BS Mapping group or sub-group, currency) — the
// same grouping the TB Pivot page already shows. Picking one of these is
// equivalent to multi-selecting every account underneath it at once, so a
// whole pivot-level group can be clubbed against a submission total in a
// single click instead of hunting down each account.
function flattenTbGroups(tree){
  const out=[];
  function walk(node,path){
    const label = node.name==='(Direct)' ? path.join(' / ') : path.concat([node.name]).join(' / ');
    const nextPath = node.name==='(Direct)' ? path : path.concat([node.name]);
    const allAccounts=collectTbAccounts(node);
    Object.entries(node.currency_totals||{}).forEach(([ccy,amt])=>{
      const accounts = ccy==='TOTAL' ? allAccounts : ccy==='FRX' ? allAccounts.filter(a=>a.currency!==state.localCcy) : allAccounts.filter(a=>a.currency===ccy);
      out.push({label:label||node.name,level:node.level,currency:ccy,forex:ccy==='FRX',amount:amt,accounts});
    });
    (node.children||[]).forEach(c=>walk(c,nextPath));
  }
  (tree||[]).forEach(g=>walk(g,[]));
  return out;
}
function filteredIndexed(list,query){
  const q=(query||'').toLowerCase().replace(/\*/g,'').trim();   // "frx*" finds the FRX rows
  const idx=list.map((it,i)=>({it,i}));
  return q?idx.filter(({it})=>JSON.stringify(it).toLowerCase().includes(q)):idx;
}
// merges whatever's checked at both the account level and the group
// (pivot) level into one deduped list of TB accounts, so a user can freely
// mix "this whole group" with "plus this one extra account" in one match
// without double-counting anything the group already covers.
// Lines picked via an FRX* row are folded into one currency:'FRX' component per
// account (and win over the same line picked as a plain USD/EUR row, so nothing
// double-counts). That FRX component is what gets saved in a mapping.
function collectSelectedTbAccounts(){
  const map=new Map(), frx=new Set();
  const add=(rows,isFrx)=>rows.forEach(a=>{ const k=rawAcctKey(a); map.set(k,a); if(isFrx) frx.add(k); });
  [...state.tbSel].map(i=>state.tbItems[i]).filter(Boolean).forEach(it=>add(it.members,it.forex));
  [...state.tbSelGroups].map(i=>state.tbGroupItems[i]).filter(Boolean).forEach(g=>add(g.accounts,g.forex));
  const out=[], fold=new Map();
  map.forEach((a,k)=>{
    if(!frx.has(k)){ out.push(a); return; }
    const fk=[a.account,a.account_desc,a.bs_mapping].join('|');
    let f=fold.get(fk);
    if(!f){ f={account:a.account,account_desc:a.account_desc,bs_mapping:a.bs_mapping,currency:'FRX',forex:true,amount:0,breakdown:{}}; fold.set(fk,f); out.push(f); }
    f.amount+=a.amount; f.breakdown[a.currency]=(f.breakdown[a.currency]||0)+a.amount;
  });
  return out;
}
function renderTbPicker(){
  const list=$('tbPickerList'); if(!list) return;
  document.querySelectorAll('#tbLevelToggle .levelBtn').forEach(b=>b.classList.toggle('active',b.dataset.level===state.tbLevel));
  document.querySelectorAll('#acctViewToggle .levelBtn').forEach(b=>b.classList.toggle('active',b.dataset.view===state.acctView));
  const isGroup=state.tbLevel==='group';
  const source=isGroup?state.tbGroupItems:state.tbItems;
  const selSet=isGroup?state.tbSelGroups:state.tbSel;
  const action=isGroup?'toggle-tb-group':'toggle-tb-item';
  $('tbPickerSearch').placeholder=isGroup?'Search BS mapping group (e.g. Assets, Other Assets)':'Search account, description or BS mapping';
  const all=filteredIndexed(source,$('tbPickerSearch')?.value);
  const shown=all.slice(0,300);
  list.innerHTML=shown.map(({it,i})=>isGroup?`
    <label class="pickerItem">
      <input type="checkbox" data-action="${action}" data-index="${i}" ${selSet.has(i)?'checked':''}>
      <div class="pMeta"><div class="pTitle">${esc(it.label)} <span class="chip">${it.accounts.length} acct${it.accounts.length!==1?'s':''}</span>${it.forex?' <span class="chip chipClub">Forex</span>':''}</div><div class="pSub">${it.forex?`All non-${esc(state.localCcy)} currencies, re-totalled each time the mapping is applied`:'BS Mapping group total (pivot level)'}</div></div>
      <div class="pAmt">${esc(ccyLabel(it.currency))} ${esc(fmt(it.amount))}</div>
    </label>`:`
    <label class="pickerItem">
      <input type="checkbox" data-action="${action}" data-index="${i}" ${selSet.has(i)?'checked':''}>
      <div class="pMeta"><div class="pTitle">${esc(aggTitle(it,state.acctView))}${it.members.length>1?` <span class="chip">${it.members.length} lines${it.accounts.length>1?' · '+it.accounts.length+' accts':''}</span>`:''}</div><div class="pSub">${esc(it.bs_mapping)}${aggAccountsNote(it,state.acctView)?' · '+esc(aggAccountsNote(it,state.acctView)):''}${it.forex?` · Forex: all non-${esc(state.localCcy)} (${esc(it.ccys.join(', '))} now)`:''}</div></div>
      <div class="pAmt">${esc(ccyLabel(it.currency))} ${esc(fmt(it.amount))}</div>
    </label>`).join('') || `<div class="muted" style="padding:14px">${isGroup?'No BS Mapping groups yet':'No Trial Balance items yet'} — upload a Trial Balance first.</div>`;
  if(all.length>300) list.insertAdjacentHTML('beforeend',`<div class="muted" style="padding:8px 11px">Showing first 300 of ${all.length} — refine your search</div>`);
  updateTbPickerSum();
}
function renderSubPicker(){
  const list=$('subPickerList'); if(!list) return;
  const all=filteredIndexed(state.subPreview,$('subPickerSearch')?.value);
  const shown=all.slice(0,300);
  list.innerHTML=shown.map(({it,i})=>`
    <label class="pickerItem">
      <input type="checkbox" data-action="toggle-sub-item" data-index="${i}" ${state.subSel.has(i)?'checked':''}>
      <div class="pMeta"><div class="pTitle">${esc(it.line_description)}${it.is_total?' <span class="chip chipTotal">TOTAL</span>':''}</div><div class="pSub">${esc(it.submission_file)} · ${esc(it.sheet)} · ${esc(it.source_cell)}</div></div>
      <div class="pAmt">${esc(it.currency)} ${esc(fmt(it.normalized_amount))}</div>
    </label>`).join('') || '<div class="muted" style="padding:14px">No submission items yet — extract sheets first.</div>';
  if(all.length>300) list.insertAdjacentHTML('beforeend',`<div class="muted" style="padding:8px 11px">Showing first 300 of ${all.length} — refine your search</div>`);
  updateSubPickerSum();
}
function updateTbPickerSum(){
  const sel=collectSelectedTbAccounts();
  const pickCount=state.tbSel.size+state.tbSelGroups.size;
  if($('tbPickerCount')) $('tbPickerCount').textContent=pickCount+' selected'+(sel.length!==pickCount?` (${sel.length} accounts)`:'');
  if($('tbPickerSum')) $('tbPickerSum').textContent=fmt(sel.reduce((a,it)=>a+it.amount,0));
  updateAmountSearchLabel();
  updateMatchPreview();
}
// combines the index-based submission picks with anything added from an
// amount search (which can reference rows beyond the first 300 shown in
// the picker's own filtered list, or even beyond the picker's search text
// entirely — the search runs over every extracted row).
function collectSelectedSubItems(){
  const picked=[...state.subSel].map(i=>state.subPreview[i]).filter(Boolean);
  return [...picked, ...state.subExtra];
}
function updateSubPickerSum(){
  const sel=collectSelectedSubItems();
  const pickCount=state.subSel.size+state.subExtra.length;
  if($('subPickerCount')) $('subPickerCount').textContent=pickCount+' selected';
  if($('subPickerSum')) $('subPickerSum').textContent=fmt(sel.reduce((a,it)=>a+it.normalized_amount,0));
  updateMatchPreview();
}

// common reporting-scale multipliers to check purely from the numbers —
// mirrors bahrain_iraq_algorithm.engine.best_scale_fit. Lets the manual
// builder catch "submission is quietly divided by 1000" even when nothing
// in the sheet's header text said so.
// 1 is included so a pure sign flip (e.g. a liability shown as -124,591,456 in the
// TB and +124,591,456 in the submission) is recognised without any scale change.
const SCALE_CANDIDATES=[1,1000,1000000,0.001,0.000001];
function fitDescription(fit){
  const parts=[]; if(fit.scale!==1) parts.push(`${scaleLabel(fit.scale)} scale`); if(fit.sign<0) parts.push('sign flipped');
  return parts.join(' and ');
}
function detectScale(tbSum,rawSubSum){
  const tolAbs=+($('tolAbs')?.value||1), tolPct=+($('tolPct')?.value||0.0001);
  const tol=v=>Math.max(tolAbs,Math.abs(v)*tolPct);
  if(Math.abs(tbSum-rawSubSum)<=tol(tbSum)) return null; // already matches at face value
  let best=null;
  for(const scale of SCALE_CANDIDATES){
    for(const sign of [1,-1]){
      const v=rawSubSum*scale*sign;
      const diff=Math.abs(tbSum-v);
      if(diff<=tol(tbSum) && (!best||diff<best.diff)) best={scale,sign,diff};
    }
  }
  return best;
}
function updateMatchPreview(){
  const box=$('matchPreview'); if(!box) return;
  const tbSel=collectSelectedTbAccounts();
  const subSel=collectSelectedSubItems();
  if(!tbSel.length||!subSel.length){ box.innerHTML=''; return; }
  const tbSum=tbSel.reduce((a,it)=>a+it.amount,0);
  const hasPresets=subSel.some(it=>it._multiplier!=null);
  const rawSubSum=hasPresets
    ? subSel.reduce((a,it)=>a+it.normalized_amount*(it._multiplier!=null?it._multiplier:1)*(it._sign!=null?it._sign:1),0)
    : subSel.reduce((a,it)=>a+it.normalized_amount,0);
  if(hasPresets){
    const tolAbs=+($('tolAbs')?.value||1), tolPct=+($('tolPct')?.value||0.0001);
    const tol=Math.max(tolAbs,Math.abs(tbSum)*tolPct);
    const isMatch=Math.abs(tbSum-rawSubSum)<=tol;
    box.innerHTML=isMatch
      ? `<div class="notice good" style="margin:0 0 14px">✓ Values match — TB ${esc(fmt(tbSum))} vs submission ${esc(fmt(rawSubSum))} (using the scale/sign found by amount search).</div>`
      : `<div class="notice" style="margin:0 0 14px">⚠ TB ${esc(fmt(tbSum))} vs submission ${esc(fmt(rawSubSum))} (after the amount-search transform) still don't agree — double-check before adding.</div>`;
    return;
  }
  const fit=detectScale(tbSum,rawSubSum);
  if(!fit){
    const tolAbs=+($('tolAbs')?.value||1), tolPct=+($('tolPct')?.value||0.0001);
    const tol=Math.max(tolAbs,Math.abs(tbSum)*tolPct);
    const isMatch=Math.abs(tbSum-rawSubSum)<=tol;
    box.innerHTML=isMatch
      ? `<div class="notice good" style="margin:0 0 14px">✓ Values match directly — TB ${esc(fmt(tbSum))} vs submission ${esc(fmt(rawSubSum))}.</div>`
      : '';
  }else{
    if(fit.scale===1 && fit.sign<0){
      box.innerHTML=`<div class="notice good" style="margin:0 0 14px">✓ Equal and opposite — TB ${esc(fmt(tbSum))} vs submission ${esc(fmt(rawSubSum))}. The submission side will be sign-flipped when you add the match.</div>`;
      return;
    }
    box.innerHTML=`<div class="notice" style="margin:0 0 14px">⚠ TB ${esc(fmt(tbSum))} vs submission ${esc(fmt(rawSubSum))} don't match directly, but with the ${esc(fitDescription(fit))} the submission becomes ${esc(fmt(rawSubSum*fit.scale*fit.sign))} — this will be applied automatically when you add the match.</div>`;
  }
}

/* ---------------- Amount search: reverse lookup from a TB club's total
   into every extracted submission line, across every sheet/file, not just
   whatever's currently filtered in the picker. ---------------- */

function updateAmountSearchLabel(){
  const lbl=$('amountSearchLabel'); if(!lbl) return;
  const tbSel=collectSelectedTbAccounts();
  if(!tbSel.length){ lbl.textContent='Select Trial Balance items first, then search for that total across every extracted sheet.'; return; }
  const sum=tbSel.reduce((a,it)=>a+it.amount,0);
  lbl.textContent=`Will search for: ${fmt(sum)} (current TB selected total)`;
}
async function searchAmount(){
  const tbSel=collectSelectedTbAccounts();
  if(!tbSel.length){toast('Select Trial Balance items first');return}
  if(!state.subPreview.length){toast('Extract submission sheets first');return}
  const amount=tbSel.reduce((a,it)=>a+it.amount,0);
  setStatus('Searching submissions for '+fmt(amount));
  const j=await api('/api/submissions/search-amount',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({session_id:state.session,amount,tolerance_abs:+$('tolAbs').value||1,tolerance_pct:+$('tolPct').value||0.0001})});
  state.amountSearchResults=j.results||[]; state.amountSearchAdded=new Set();
  state.amountSearchTarget=amount; state.amountSearchTotalFound=j.total_found||0;
  renderAmountSearchResults();
  setStatus(state.amountSearchResults.length?`Found ${j.total_found} match(es)`:'No matches found');
}
function renderAmountSearchResults(){
  const box=$('amountSearchResults'); if(!box) return;
  if(!state.amountSearchResults.length){ box.innerHTML=''; return; }
  const target=state.amountSearchTarget;
  box.innerHTML=`<div class="notice good" style="margin-bottom:8px">Found ${state.amountSearchTotalFound} line(s) across all extracted sheets matching ${esc(fmt(target))}${state.amountSearchTotalFound>200?' (showing the closest 200)':''}.</div>`+
    state.amountSearchResults.map((r,i)=>{
      const added=state.amountSearchAdded.has(i);
      const transform=[];
      if(r.scale!==1) transform.push(scaleLabel(r.scale));
      if(r.sign<0) transform.push('sign flipped');
      const transformLabel=transform.length?transform.join(', '):'exact match';
      return `<div class="searchResultRow ${added?'added':''}">
        <div class="srMeta">
          <div class="srTitle">${esc(r.line_description)}${r.is_total?' <span class="chip chipTotal">TOTAL</span>':''}</div>
          <div class="srSub">${esc(r.submission_file)} · ${esc(r.sheet)} · ${esc(r.source_cell)} · ${esc(r.currency)} · ${esc(transformLabel)}</div>
        </div>
        <div class="srAmt">${esc(fmt(r.matched_value))}<small>sheet shows ${esc(fmt(r.raw_amount))}</small></div>
        <button type="button" class="secondary" data-action="add-search-result" data-index="${i}" ${added?'disabled':''}>${added?'Added ✓':'+ Add'}</button>
      </div>`;
    }).join('');
}
function addSearchResultToMatch(i){
  const r=state.amountSearchResults[i]; if(!r) return;
  state.amountSearchAdded.add(i);
  state.subExtra.push({
    submission_file:r.submission_file,sheet:r.sheet,row_number:r.row_number,source_cell:r.source_cell,
    line_description:r.line_description,currency:r.currency,is_total:r.is_total,
    normalized_amount:r.raw_amount,_multiplier:r.scale,_sign:r.sign
  });
  renderAmountSearchResults(); updateSubPickerSum();
  toast('Added to submission side of the match');
}
function addManualMatch(){
  const tbSelItems=collectSelectedTbAccounts();
  const subSelItems=collectSelectedSubItems();
  if(!tbSelItems.length||!subSelItems.length){toast('Select at least one item on each side');return}
  const label=($('matchLabel').value||'').trim();
  if(!label){toast('Give this match a short description');return}
  const ccySet=new Set(tbSelItems.map(it=>it.currency));
  const currency=ccySet.size===1?[...ccySet][0]:'TOTAL';
  if(ccySet.size>1) toast('Selected TB items span multiple currencies — compared as a combined TOTAL.');
  const tbSum=tbSelItems.reduce((a,it)=>a+it.amount,0);
  // items added via amount search already carry the exact scale/sign that
  // was found to match — trust those per-item rather than re-detecting a
  // single uniform scale across everything selected, which would be wrong
  // if a search-matched item is mixed in with plain, unscaled picks.
  const hasPresets=subSelItems.some(it=>it._multiplier!=null);
  let scale=1, signAdj=1;
  if(!hasPresets){
    const rawSubSum=subSelItems.reduce((a,it)=>a+it.normalized_amount,0);
    const fit=detectScale(tbSum,rawSubSum);
    if(fit){ toast(`Detected ${fit.scale===1?'opposite signs':'a '+fitDescription(fit)} — applied automatically.`); scale=fit.scale; signAdj=fit.sign; }
  }
  const tb_components=tbSelItems.map(it=>({account:it.account,account_desc:it.account_desc,bs_mapping:it.bs_mapping,currency:it.currency,amount:it.amount,sign:1,...(it.breakdown?{breakdown:it.breakdown}:{})}));
  const components=subSelItems.map(it=>({submission_file:it.submission_file,sheet:it.sheet,row_number:it.row_number,source_cell:it.source_cell,
    line_description:it.line_description,currency:it.currency,is_total:!!it.is_total,amount:it.normalized_amount,
    multiplier:it._multiplier!=null?it._multiplier:scale,sign:it._sign!=null?it._sign:signAdj}));
  state.manualMatches.push({bs_mapping:label,label,currency,rule_type:'MANUAL_CLUB',source:'MANUAL',tb_components,components,warnings:[]});
  state.tbSel.clear(); state.tbSelGroups.clear(); state.subSel.clear(); state.subExtra=[]; state.amountSearchAdded=new Set();
  state.amountSearchResults=[]; $('matchLabel').value='';
  renderTbPicker(); renderSubPicker(); renderManualMatches(); renderAmountSearchResults(); updateMatchPreview();
  toast('Manual match added');
}
function renderManualMatches(){
  $('manualMatches').innerHTML=state.manualMatches.map((m,i)=>{
    const tbSum=(m.tb_components||[]).reduce((a,c)=>a+c.amount*(c.sign||1),0);
    const subSum=(m.components||[]).reduce((a,c)=>a+c.amount*(c.sign||1)*(c.multiplier||1),0);
    const diff=tbSum-subSum;
    const noEvidence=!(m.tb_components||[]).length && !(m.components||[]).length;
    const status=noEvidence?'UNRESOLVED':(Math.abs(diff)<1e-6?'MATCH':(Math.abs(diff)<=1||Math.abs(diff)/Math.max(Math.abs(tbSum),1)<=0.0001?'MATCH_WITHIN_TOLERANCE':'REVIEW_REQUIRED'));
    const subScales=new Set((m.components||[]).map(c=>c.multiplier||1));
    const appliedScale=subScales.size===1?[...subScales][0]:null;
    const sLabel=scaleLabel(appliedScale);
    const tbAgg=aggregateAccounts(m.tb_components||[],state.acctView,c=>c.amount*(c.sign||1));
    return `<div class="manualMatchCard ${(m.warnings&&m.warnings.length)?'hasWarning':''}">
      <div class="manualMatchHead">
        <b>${esc(m.label||m.bs_mapping)}</b>
        <span class="chip">${esc(ccyLabel(m.currency))}</span>
        <span class="chip">${esc(m.source||'MANUAL')}</span>
        ${sLabel?`<span class="chip chipClub">Scale ${esc(sLabel)} applied</span>`:''}
        ${(m.components||[]).length&&(m.components||[]).every(c=>(c.sign||1)<0)?'<span class="chip chipClub" title="Submission values are equal and opposite to the Trial Balance (e.g. liabilities), so they were sign-flipped to reconcile">Sign flipped</span>':''}
        <span class="pill ${status}">${status}</span>
        <button type="button" class="danger" data-action="remove-manual-match" data-index="${i}">Remove</button>
      </div>
      <div class="matchSides">
        <div class="matchSide"><small>Trial Balance (${tbAgg.length}${tbAgg.length!==(m.tb_components||[]).length?' by '+(state.acctView==='no'?'account no.':'description')+' · '+(m.tb_components||[]).length+' lines':''})</small>
          <ul>${tbAgg.map(g=>{ const bd={}; g.members.forEach(c=>Object.entries(c.breakdown||{}).forEach(([k,v])=>{bd[k]=(bd[k]||0)+v})); const bdTxt=Object.entries(bd).map(([k,v])=>`${k} ${fmt(v)}`).join(' + ');
          return `<li title="${esc(g.accounts.join(', '))}${bdTxt?esc(' — FRX* = '+bdTxt):''}"><span>${g.currency==='FRX'?'<span class="chip chipClub">FRX*</span> ':''}${esc(aggTitle(g,state.acctView))}${g.members.length>1?` <span class="muted">(${g.members.length} lines${g.accounts.length>1?', '+g.accounts.length+' accts':''})</span>`:''}</span><b>${esc(fmt(g.amount))}</b></li>`}).join('')||'<li class="muted">none resolved</li>'}</ul>
          <div class="muted" style="margin-top:6px">Total: <b>${esc(fmt(tbSum))}</b></div>
        </div>
        <div class="matchArrow">=</div>
        <div class="matchSide"><small>Submission (${(m.components||[]).length})</small>
          <ul>${(m.components||[]).map(c=>`<li><span>${esc(c.sheet)} · ${esc(c.line_description)}</span><b>${esc(fmt(c.amount))}</b></li>`).join('')||'<li class="muted">none resolved</li>'}</ul>
          <div class="muted" style="margin-top:6px">Total: <b>${esc(fmt(subSum))}</b></div>
        </div>
      </div>
      ${noEvidence?'':`<div class="matchVariance"><span class="muted">Variance (TB - submission):</span> <b class="${status==='REVIEW_REQUIRED'?'bad':'good'}">${esc(fmt(diff))}</b>${Math.abs(tbSum)>1e-9?` <span class="muted">(${esc(dmPct(diff/Math.abs(tbSum)*100))})</span>`:''}</div>`}
      ${m.warnings&&m.warnings.length?`<div class="matchWarnings">${m.warnings.map(w=>`<div>⚠ ${esc(w)}</div>`).join('')}</div>`:''}
    </div>`}).join('') || '<div class="notice">No manual matches yet — build one above or load a saved mapping file.</div>';
}

/* ---------------- Mapping Studio: save / load mapping file ---------------- */

function downloadMapping(){
  if(!state.manualMatches.length){toast('No manual matches to save yet');return}
  const matches=state.manualMatches.map(m=>({
    label:m.label||m.bs_mapping,currency:m.currency,rule_type:m.rule_type||'MANUAL_CLUB',
    tb_side:(m.tb_components||[]).map(c=>({account:c.account,account_desc:c.account_desc,bs_mapping:c.bs_mapping,currency:c.currency,sign:c.sign||1})),
    submission_side:(m.components||[]).map(c=>({match_text:c.line_description,sheet:c.sheet,currency:c.currency,is_total:!!c.is_total,sign:c.sign||1}))
  }));
  const blob=new Blob([JSON.stringify({format:'bahrain-iraq-recon-mapping',version:1,matches},null,2)],{type:'application/json'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='reconciliation_mapping.json';
  document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(a.href);
}
async function applyMapping(){
  const f=$('mappingFile').files[0];
  if(!f){toast('Choose a mapping file first');return}
  let parsed;
  try{ parsed=JSON.parse(await f.text()); }
  catch(e){ throw new Error('That file is not valid JSON.'); }
  const templates=parsed.matches||parsed.templates||(Array.isArray(parsed)?parsed:null);
  if(!templates||!templates.length) throw new Error('No matches found in that mapping file.');
  const j=await api('/api/rules/resolve',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({session_id:state.session,templates})});
  let warnCount=0;
  (j.results||[]).forEach(r=>{
    state.manualMatches.push({...r.resolved,warnings:r.warnings||[]});
    if(r.warnings&&r.warnings.length) warnCount++;
  });
  renderManualMatches();
  $('mappingImportMessage').innerHTML=`<div class="notice ${warnCount?'':'good'}">Imported ${(j.results||[]).length} match(es)${warnCount?`, ${warnCount} need review (see the warning badges below)`:''}.</div>`;
  $('mappingFile').value='';
}

/* ---------------- Mapping Coverage ---------------- */

function renderCoverage(){
  const box=$('coverageSummary'); if(!box) return;
  const cov=state.mappingCoverage||[];
  const s=state.mappingCoverageSummary;
  const reconciled=cov.filter(e=>e.reconciled_status==='MATCH'||e.reconciled_status==='MATCH_WITHIN_TOLERANCE').length;
  const needsReview=cov.filter(e=>e.reconciled_status==='REVIEW_REQUIRED'||e.reconciled_status==='UNRESOLVED').length;
  const notApproved=cov.filter(e=>e.reconciled_status==='NOT_APPROVED').length;
  box.innerHTML=[
    ['Total rules',s?s.total:cov.length],
    ['Fulfilled (data found)',s?s.fulfilled:cov.filter(e=>e.fulfilled).length],
    ['Unresolved (no data)',s?s.unresolved:cov.filter(e=>!e.fulfilled).length],
    ['Reconciled (matched)',reconciled],
    ['Needs review',needsReview],
  ].map(([k,v])=>`<div class="kpi"><small>${esc(k)}</small><b>${esc(v)}</b></div>`).join('');
  renderCoverageTable();
}
function renderCoverageTable(){
  const t=$('coverageTable'); if(!t) return;
  const q=($('coverageSearch')?.value||'').toLowerCase();
  const f=state.coverageFilter||'all';
  let rows=state.mappingCoverage||[];
  if(f==='fulfilled') rows=rows.filter(e=>e.fulfilled);
  else if(f==='unresolved') rows=rows.filter(e=>!e.fulfilled);
  else if(f==='reconciled') rows=rows.filter(e=>e.reconciled_status==='MATCH'||e.reconciled_status==='MATCH_WITHIN_TOLERANCE');
  else if(f==='review') rows=rows.filter(e=>e.reconciled_status==='REVIEW_REQUIRED'||e.reconciled_status==='UNRESOLVED');
  if(q) rows=rows.filter(e=>JSON.stringify(e).toLowerCase().includes(q));
  table('coverageTable',rows,[
    {key:'label',label:'Mapping rule'},
    {key:'currency',label:'Currency'},
    {key:'fulfilled',label:'Data found?',render:v=>v?'<span class="pill MATCH">FULFILLED</span>':'<span class="pill REVIEW_REQUIRED">NOT FOUND</span>'},
    {key:'tb_accounts',label:'TB accounts'},
    {key:'sub_lines',label:'Submission lines'},
    {key:'reconciled_status',label:'Reconciliation outcome',render:v=>v?`<span class="pill ${v==='NOT_APPROVED'?'NO_RULE':v}">${esc(v)}</span>`:'<span class="muted">not run</span>'},
    {key:'tb_amount',label:'TB Amount'},{key:'submission_amount',label:'Submission Amount'},{key:'difference',label:'Difference'},
    {key:'warnings',label:'Detail',render:v=>(v&&v.length)?esc(v.join(' · ')):''}
  ]);
}

async function runRecon(){
  const approved=state.suggestions.filter((_,i)=>state.suggestionDecisions[i]).map(s=>({bs_mapping:s.bs_mapping,currency:s.currency,rule_type:s.rule_type,source:'AUTO',components:s.components}));
  const manual=state.manualMatches.map(m=>({bs_mapping:m.label||m.bs_mapping,currency:m.currency,rule_type:m.rule_type,source:m.source||'MANUAL',tb_components:m.tb_components,components:m.components}));
  const rules=[...approved,...manual];
  if(!rules.length){toast('Approve at least one suggestion or build a manual match');return}
  const j=await api('/api/reconcile',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({session_id:state.session,rules,tolerance_abs:+$('tolAbs').value,tolerance_pct:+$('tolPct').value})});
  state.results=j.results||[]; state.lineage=j.lineage||[];
  if(j.mapping_coverage) state.mappingCoverage=j.mapping_coverage;
  renderCoverage();
  let counts={}; state.results.forEach(x=>counts[x.status]=(counts[x.status]||0)+1);
  $('resultSummary').innerHTML=Object.entries(counts).map(([k,v])=>`<div class="kpi"><small>${esc(k)}</small><b>${esc(v)}</b></div>`).join('');
  table('resultTable',state.results,[
    {key:'bs_mapping',label:'BS Mapping'},{key:'currency',label:'Currency'},
    {key:'tb_amount',label:'TB Amount'},{key:'submission_amount',label:'Submission Amount'},{key:'difference',label:'Difference'},
    {key:'variance_pct',label:'Variance %',render:v=>v==null?'—':(v*100).toFixed(4)+'%'},
    {key:'status',label:'Status',render:v=>`<span class="pill ${v}">${esc(v)}</span>`},
    {key:'rule_type',label:'Rule'},{key:'source',label:'Source'}]);
  renderLineageTree();
  setStatus('Reconciliation complete'); go('results');
}

/* ---------------- Reconciliation Map ---------------- */

function statusPill(status){return status?`<span class="pill ${status}">${esc(status)}</span>`:''}
function renderMatchRow(m){
  let evidence=[]; try{evidence=JSON.parse(m.evidence||'[]')}catch(e){}
  let tbEvidence=[]; try{tbEvidence=JSON.parse(m.tb_evidence||'[]')}catch(e){}
  return `<details class="matchRow">
    <summary><span class="chip">${esc(m.currency)}</span>${statusPill(m.status)}
      <span class="muted">TB ${esc(fmt(m.tb_amount))} vs submission ${esc(fmt(m.submission_amount))} (Δ ${esc(fmt(m.difference))})</span>
      <span class="chip">${esc(m.source||'')}</span></summary>
    ${tbEvidence.length?`<div class="muted" style="margin:8px 0 2px">Trial Balance side</div>
    <table class="miniTable"><thead><tr><th>Account</th><th>Description</th><th>BS Mapping</th><th>Amount</th><th>Currency</th></tr></thead>
    <tbody>${tbEvidence.map(e=>`<tr><td>${esc(e.account)}</td><td>${esc(e.description)}</td><td>${esc(e.bs_mapping)}</td><td>${esc(fmt(e.amount))}</td><td>${esc(e.currency||'')}</td></tr>`).join('')}</tbody></table>
    <div class="muted" style="margin:10px 0 2px">Submission side</div>`:''}
    <table class="miniTable"><thead><tr><th>File</th><th>Sheet</th><th>Row</th><th>Description</th><th>Amount</th><th>Currency</th></tr></thead>
    <tbody>${evidence.map(e=>`<tr><td>${esc(e.file)}</td><td>${esc(e.sheet)}</td><td>${esc(e.row)}</td><td>${esc(e.description)}</td><td>${esc(fmt(e.amount))}</td><td>${esc(e.currency||'')}</td></tr>`).join('')||'<tr><td colspan="6" class="muted">No evidence rows</td></tr>'}</tbody></table>
  </details>`;
}
function renderLineageNode(node){
  const matches=node.matches||[];
  const children=node.children||[];
  const accounts=node.accounts||[];
  const cls=node.level==='group'?'groupNode':'subNode';
  return `<details class="treeNode ${cls}" open>
    <summary><span class="nodeName">${esc(node.name)}</span>${currencyChips(node.currency_totals)}${matches.length?`<span class="chip chipMatches">${matches.length} match${matches.length!==1?'es':''}</span>`:''}</summary>
    <div class="treeChildren">
      ${matches.length?`<div class="matchList">${matches.map(renderMatchRow).join('')}</div>`:(children.length?'':'<div class="muted" style="padding:8px 4px">No reconciliation rule approved for this line yet.</div>')}
      ${accounts.length?`<details class="accountsDetail"><summary class="muted">${accounts.length} TB account row(s)</summary><div class="accountList">${accounts.map(a=>`<div class="acctRow"><span>${esc(a.account)} — ${esc(a.account_desc)}</span><span class="muted">${esc(a.currency)}</span><b>${esc(fmt(a.amount))}</b></div>`).join('')}</div></details>`:''}
      ${children.map(renderLineageNode).join('')}
    </div>
  </details>`;
}
function collectAttachedKeys(tree){
  const keys=new Set();
  (function walk(nodes){
    (nodes||[]).forEach(n=>{ (n.matches||[]).forEach(m=>keys.add(m.bs_mapping+'|'+m.currency)); walk(n.children); });
  })(tree);
  return keys;
}
function renderLineageTree(){
  $('mapLegend').innerHTML=['MATCH','MATCH_WITHIN_TOLERANCE','REVIEW_REQUIRED','UNRESOLVED'].map(s=>`<span class="pill ${s}">${s}</span>`).join(' ');
  const treeHtml=(state.lineage||[]).map(renderLineageNode).join('');
  // manual / imported matches carry their own free-text label rather than a
  // TB group name, so they never line up with a tree node above — list them
  // in their own section instead of letting them silently vanish from the map.
  const attached=collectAttachedKeys(state.lineage);
  const orphans=(state.results||[]).filter(r=>!attached.has(r.bs_mapping+'|'+r.currency));
  const orphanHtml=orphans.length?`<details class="treeNode groupNode" open>
      <summary><span class="nodeName">Manual &amp; imported matches</span><span class="chip chipMatches">${orphans.length} match${orphans.length!==1?'es':''}</span></summary>
      <div class="treeChildren"><div class="matchList">${orphans.map(renderMatchRow).join('')}</div></div>
    </details>`:'';
  $('lineageTree').innerHTML=(treeHtml+orphanHtml) || '<div class="muted">Run the reconciliation to build the map.</div>';
}

async function downloadOutput(){
  if(!state.session){toast('No active session yet');return}
  await downloadFile('/api/download?session_id='+encodeURIComponent(state.session),'Bahrain_Iraq_Reconciliation_Output.xlsx');
}

/* ---------------- event wiring (delegated — no inline handlers) ---------------- */

const CLICK_ACTIONS={
  'go':(el)=>go(el.dataset.target),
  'choose-tb':()=>$('tbFile').click(),
  'choose-subs':()=>$('subFiles').click(),
  'choose-mapping':()=>$('mappingFile').click(),
  'upload-tb':(el)=>guarded(el,uploadTB),
  'upload-subs':(el)=>guarded(el,uploadSubs),
  'extract-subs':(el)=>guarded(el,extractSubs),
  'download-submissions':(el)=>guarded(el,downloadSubmissions),
  'run-depth':(el)=>guarded(el,runDepth),
  'dm-rerun':(el)=>guarded(el,rerunDefault),
  'dm-pdf':(el)=>guarded(el,downloadDefaultPdf),
  'dm-download':(el)=>guarded(el,()=>downloadFile('/api/default/download?session_id='+encodeURIComponent(state.session)+'&file='+encodeURIComponent(el.dataset.file),'DefaultMapping.xlsx')),
  'dm-download-all':(el)=>guarded(el,()=>downloadFile('/api/default/download-all?session_id='+encodeURIComponent(state.session),'Default_Mapping_Outputs.zip')),
  'set-lang':(el)=>setLang(el.dataset.lang),
  'choose-dict':()=>$('dictFile').click(),
  'proceed-lang':(el)=>guarded(el,proceedLang),
  'download-missing':(el)=>guarded(el,()=>downloadFile('/api/submissions/translation-download?session_id='+encodeURIComponent(state.session)+'&kind=missing','Missing_Arabic_Words.xlsx')),
  'download-translated':(el)=>guarded(el,()=>downloadFile('/api/submissions/translation-download?session_id='+encodeURIComponent(state.session)+'&kind=file&file='+encodeURIComponent(el.dataset.file),'translated.xlsx')),
  'choose-ob':()=>$('obFile').click(),
  'upload-ob':(el)=>guarded(el,uploadOutstanding),
  'download-outstanding':(el)=>guarded(el,()=>downloadFile('/api/outstanding/download?session_id='+encodeURIComponent(state.session),'Outstanding_Report_IRAQ_Pivot_Output.xlsx')),
  'download-depth':(el)=>guarded(el,()=>downloadFile('/api/depth-search/download?session_id='+encodeURIComponent(state.session)+'&file='+encodeURIComponent(el.dataset.file),'DepthSearch.xlsx')),
  'download-depth-all':(el)=>guarded(el,()=>downloadFile('/api/depth-search/download-all?session_id='+encodeURIComponent(state.session),'Depth_Search_Outputs.zip')),
  'set-tb-level':(el)=>{ state.tbLevel=el.dataset.level; renderTbPicker(); },
  'set-acct-view':(el)=>{
    if(el.dataset.view===state.acctView) return;
    // keep a pick only if every raw line behind it is still fully covered after re-aggregating
    // FRX* picks and plain-currency picks are tracked apart so each comes back as itself
    const pickedRaw=new Set(), pickedFrx=new Set();
    [...state.tbSel].forEach(i=>{ const it=state.tbItems[i]; if(it) it.members.forEach(a=>(it.forex?pickedFrx:pickedRaw).add(rawAcctKey(a))); });
    state.acctView=el.dataset.view;
    state.tbItems=aggregateAccounts(state.tbRaw,state.acctView,null,true);
    const before=pickedRaw.size+pickedFrx.size; state.tbSel=new Set();
    state.tbItems.forEach((it,i)=>{ const src=it.forex?pickedFrx:pickedRaw; if(it.members.every(a=>src.has(rawAcctKey(a)))) state.tbSel.add(i); });
    const kept=new Set(); [...state.tbSel].forEach(i=>{ const it=state.tbItems[i]; it.members.forEach(a=>kept.add((it.forex?'F:':'P:')+rawAcctKey(a))); });
    if(kept.size<before) toast('Some picks only partly covered the new grouping and were dropped — re-pick them.');
    renderTbPicker(); renderManualMatches();
  },
  'select-all-tb':()=>{
    const source=state.tbLevel==='group'?state.tbGroupItems:state.tbItems;
    const selSet=state.tbLevel==='group'?state.tbSelGroups:state.tbSel;
    filteredIndexed(source,$('tbPickerSearch')?.value).forEach(({i})=>selSet.add(i));
    renderTbPicker();
  },
  'clear-tb':()=>{
    const source=state.tbLevel==='group'?state.tbGroupItems:state.tbItems;
    const selSet=state.tbLevel==='group'?state.tbSelGroups:state.tbSel;
    filteredIndexed(source,$('tbPickerSearch')?.value).forEach(({i})=>selSet.delete(i));
    renderTbPicker();
  },
  'select-all-sub':()=>{ filteredIndexed(state.subPreview,$('subPickerSearch')?.value).forEach(({i})=>state.subSel.add(i)); renderSubPicker(); },
  'clear-sub':()=>{ filteredIndexed(state.subPreview,$('subPickerSearch')?.value).forEach(({i})=>state.subSel.delete(i)); renderSubPicker(); },
  'add-manual-match':()=>addManualMatch(),
  'search-amount':(el)=>guarded(el,searchAmount),
  'add-search-result':(el)=>addSearchResultToMatch(+el.dataset.index),
  'remove-manual-match':(el)=>{state.manualMatches.splice(+el.dataset.index,1);renderManualMatches()},
  'download-mapping':()=>downloadMapping(),
  'apply-mapping':(el)=>guarded(el,applyMapping),
  'run-recon':(el)=>guarded(el,runRecon),
  'download':(el)=>guarded(el,downloadOutput),
  'coverage-filter':(el)=>{
    state.coverageFilter=el.dataset.filter;
    document.querySelectorAll('[data-action="coverage-filter"]').forEach(b=>b.classList.toggle('active',b.dataset.filter===state.coverageFilter));
    renderCoverageTable();
  },
};

document.addEventListener('click',e=>{
  const nav=e.target.closest('.step');
  if(nav && nav.dataset.page){ go(nav.dataset.page); return; }
  const el=e.target.closest('[data-action]');
  if(!el) return;
  const handler=CLICK_ACTIONS[el.dataset.action];
  if(!handler) return;
  e.preventDefault();
  handler(el,e);
});

document.addEventListener('change',e=>{
  const el=e.target;
  if(el.id==='dmStatus'){ renderDefault(); return; }
  if(el.dataset.action==='toggle-sheet'){
    state.selection[el.dataset.file][el.dataset.sheet].checked=el.checked; return;
  }
  if(el.dataset.action==='set-header-row'){
    state.selection[el.dataset.file][el.dataset.sheet].header_row=el.value?parseInt(el.value,10):null; return;
  }
  if(el.dataset.action==='work-lang'){ $('dictBox').style.display=el.value==='en'?'block':'none'; return; }
  if(el.id==='dictFile'){ $('dictName').textContent=el.files[0]?el.files[0].name:''; return; }
  if(el.dataset.action==='ob-category'){ guarded(null,()=>changeObCategory(el.value)); return; }
  if(el.dataset.action==='set-file-type'){
    state.fileTypes[el.dataset.file]=el.value;
    renderDepthFiles();                       // refresh the Off-balance note above
    if(state.depth) toast('File type changed — run the depth search again to apply it.');
    return;
  }
  if(el.dataset.action==='toggle-suggestion'){
    state.suggestionDecisions[+el.dataset.index]=el.checked;
    el.closest('.suggestionCard')?.classList.toggle('accepted',el.checked);
    renderSuggestionSummary(); return;
  }
  if(el.dataset.action==='toggle-tb-item'){
    const i=+el.dataset.index; if(el.checked) state.tbSel.add(i); else state.tbSel.delete(i); updateTbPickerSum(); return;
  }
  if(el.dataset.action==='toggle-tb-group'){
    const i=+el.dataset.index; if(el.checked) state.tbSelGroups.add(i); else state.tbSelGroups.delete(i); updateTbPickerSum(); return;
  }
  if(el.dataset.action==='toggle-sub-item'){
    const i=+el.dataset.index; if(el.checked) state.subSel.add(i); else state.subSel.delete(i); updateSubPickerSum(); return;
  }
});

document.addEventListener('input',e=>{
  if(e.target.id==='pivotSearch') renderPivot();
  if(e.target.id==='tbPickerSearch') renderTbPicker();
  if(e.target.id==='subPickerSearch') renderSubPicker();
  if(e.target.id==='coverageSearch') renderCoverageTable();
  if(e.target.id==='depthFilter') renderDepthResults();
  if(e.target.id==='dmFilter') renderDefault();
});

/* ---------------- Interface language: English <-> Arabic ----------------
   The app is written in English. In Arabic mode every piece of interface text (static or rendered later by
   any table, toast or notice) is translated from this dictionary by a DOM translator, and the layout flips to
   right-to-left. Data (account names, file names, amounts, sheet names) is never translated: table cells
   and picker rows are skipped unless they hold interface text (data-tr, pills, chips). Anything not in the
   dictionary simply stays English. */

const AR_DICT={
"Recon Studio": "استوديو التسوية",
"Bahrain - Iraq": "البحرين - العراق",
"Overview": "نظرة عامة",
"Trial Balance": "ميزان المراجعة",
"TB Pivot": "الجدول المحوري للميزان",
"Submissions": "التقارير المقدّمة",
"Off-Balance": "خارج الميزانية",
"Depth Search": "البحث المتعمق",
"Mapping Studio": "استوديو الربط",
"Mapping Coverage": "تغطية الربط",
"Recon Results": "نتائج التسوية",
"Reconciliation Map": "خريطة التسوية",
"Export": "التصدير",
"Starting...": "جارٍ البدء...",
"Starting…": "جارٍ البدء…",
"Language": "اللغة",
"CONTROLLED RECONCILIATION WORKFLOW": "مسار تسوية مضبوط",
"From Trial Balance to a fully explainable reconciliation.": "من ميزان المراجعة إلى تسوية قابلة للتفسير بالكامل.",
"Build the TB pivot, choose exactly which submission sheets to read, let the engine detect currency columns, totals and section hierarchy even in imperfect layouts, then review every Assets↔Assets, sub-group↔sub-group and USD↔USD / Total↔Total match before it's approved.": "أنشئ الجدول المحوري لميزان المراجعة، واختر بدقة أوراق التقارير المقدّمة التي ستُقرأ، ودع المحرّك يكتشف أعمدة العملات والإجماليات وتسلسل الأقسام حتى في التنسيقات غير المنتظمة، ثم راجع كل مطابقة بين الأصول والأصول، والمجموعات الفرعية والمجموعات الفرعية، والدولار والدولار / الإجمالي والإجمالي قبل اعتمادها.",
"Start with Trial Balance": "ابدأ بميزان المراجعة",
"Process": "المراحل",
"Pivot": "الجدول المحوري",
"Select submission sheets": "اختيار أوراق التقارير",
"Structural extraction": "الاستخراج الهيكلي",
"Auto-match groups, sub-groups, currencies & totals": "مطابقة تلقائية للمجموعات والمجموعات الفرعية والعملات والإجماليات",
"Review / approve": "المراجعة / الاعتماد",
"Reconciliation": "التسوية",
"Full lineage map": "خريطة التتبع الكاملة",
"Excel export": "تصدير إلى Excel",
"Table-aware reading": "قراءة واعية بالجداول",
"Header, currency and total columns are detected per sheet from real formatting (bold, indentation) - not a fixed layout.": "تُكتشف أعمدة العناوين والعملات والإجماليات في كل ورقة من التنسيق الفعلي (الخط العريض والإزاحة) وليس من تخطيط ثابت.",
"Currency & total aware": "مراعاة العملة والإجمالي",
"USD is only ever compared with USD, IQD with IQD, and Total with Total. Credit/debit sign convention is inferred automatically.": "تُقارَن العملة بنفس العملة فقط: الدولار بالدولار والدينار بالدينار والإجمالي بالإجمالي. ويُستنتج اصطلاح إشارة الدائن/المدين تلقائيًا.",
"Full lineage": "تتبع كامل",
"Every accepted match keeps its file, sheet, cell and grouping evidence, shown end-to-end in the Reconciliation Map.": "تحتفظ كل مطابقة معتمدة بأدلة الملف والورقة والخلية والتجميع، وتُعرض من البداية إلى النهاية في خريطة التسوية.",
"Step 2": "الخطوة 2",
"Step 4": "الخطوة 4",
"Upload Trial Balance": "رفع ميزان المراجعة",
"Dynamic sheet and header detection. No managed folder required.": "اكتشاف ديناميكي للورقة وصف العناوين. لا حاجة إلى مجلد مُدار.",
"Drop the Trial Balance here": "أفلت ميزان المراجعة هنا",
"Excel workbook, processed temporarily in the WebApp session": "مصنف Excel تتم معالجته مؤقتًا داخل جلسة التطبيق",
"Choose file": "اختيار ملف",
"Inspect and build pivot": "فحص وبناء الجدول المحوري",
"Trial Balance Pivot": "الجدول المحوري لميزان المراجعة",
"Rows: BS Mapping → Account → Account Desc | Columns: Tran Ccy | Values: Sum of Adjusted Balance": "الصفوف: ربط الميزانية ← الحساب ← وصف الحساب | الأعمدة: عملة العملية | القيم: مجموع الرصيد المعدّل",
"Download Excel": "تنزيل Excel",
"Search mapping, account or description": "ابحث عن ربط أو حساب أو وصف",
"Group / Sub-group structure detected in the Trial Balance": "هيكل المجموعات / المجموعات الفرعية المكتشف في ميزان المراجعة",
"Used to auto-match against submission sections": "يُستخدم للمطابقة التلقائية مع أقسام التقارير المقدّمة",
"Detected sheet": "تم اكتشاف الورقة",
", header row": "، صف العناوين",
"in": "في",
"rows": "الصفوف",
"groups": "المجموعات",
"currencies": "العملات",
"balance": "الرصيد",
"unmapped": "غير مربوط",
"No structure detected yet.": "لم يُكتشف أي هيكل بعد.",
"No account rows": "لا توجد صفوف حسابات",
"Upload Submission Files": "رفع ملفات التقارير المقدّمة",
"Upload up to eight workbooks, then choose exactly which sheets to read into the reconciliation.": "ارفع حتى ثمانية مصنفات، ثم اختر بدقة الأوراق التي ستُقرأ ضمن التسوية.",
"Select the submissions": "اختيار التقارير المقدّمة",
"Assets, liabilities/capital, and other supporting submissions (max 8)": "الأصول والالتزامات/رأس المال وسائر التقارير الداعمة (بحد أقصى 8)",
"Choose files": "اختيار ملفات",
"Inspect workbooks": "فحص المصنفات",
"Arabic content detected": "تم اكتشاف محتوى عربي",
"Choose the language you want to work in, then press Proceed.": "اختر اللغة التي تريد العمل بها، ثم اضغط «متابعة».",
"English": "الإنجليزية",
"Arabic": "العربية",
"Translate the submissions to English using your own dictionary, then work in English.": "ترجم التقارير المقدّمة إلى الإنجليزية باستخدام قاموسك الخاص، ثم اعمل بالإنجليزية.",
"Keep the submissions in Arabic and work on them as they are.": "أبقِ التقارير المقدّمة بالعربية واعمل عليها كما هي.",
"Dictionary workbook (Arabic in column A, English in column B, header in row 1)": "مصنف القاموس (العربية في العمود A والإنجليزية في العمود B والعناوين في الصف 1)",
"Choose dictionary": "اختيار القاموس",
"Proceed": "متابعة",
"Extract selected sheets": "استخراج الأوراق المحددة",
"Extracted lines": "البنود المستخرجة",
"The cleaned, currency-tagged lines read from your selected sheets": "البنود المنظّفة والموسومة بالعملة المقروءة من الأوراق التي اخترتها",
"Download submissions file": "تنزيل ملف التقارير المقدّمة",
"Header row": "صف العناوين",
"No currency columns detected - will use raw numeric cells": "لم تُكتشف أعمدة عملات - ستُستخدم الخلايا الرقمية الخام",
"Working in Arabic - the submissions are used exactly as uploaded. Choose the sheets to read, then extract.": "العمل بالعربية - تُستخدم التقارير المقدّمة كما رُفعت تمامًا. اختر الأوراق المراد قراءتها ثم نفّذ الاستخراج.",
"Download missing terms": "تنزيل المصطلحات الناقصة",
"No submission files yet - upload them in the Submissions step.": "لا توجد ملفات تقارير بعد - ارفعها في خطوة التقارير المقدّمة.",
"Off-balance": "خارج الميزانية",
"Outstanding Report": "تقرير الأرصدة القائمة",
"Upload the Outstanding Report. The raw-data sheet and header row are found automatically, Equ-IQD is cleaned and validated, and two pivots are built. Depth Search then looks for their values in the Off-balance (064) submission, with the Bucket pivot searched on its Maturity sheet only.": "ارفع تقرير الأرصدة القائمة. تُكتشف ورقة البيانات الخام وصف العناوين تلقائيًا، وتُنظَّف قيمة المكافئ بالدينار العراقي وتُراجع، ويُبنى جدولان محوريان. ثم يبحث البحث المتعمق عن قيمهما في تقرير خارج الميزانية (064)، ويُبحث في الجدول المحوري للفترات في ورقة الاستحقاق فقط.",
"Drop the Outstanding Report here": "أفلت تقرير الأرصدة القائمة هنا",
"Excel workbook with the raw outstanding-report data": "مصنف Excel يحتوي على البيانات الخام لتقرير الأرصدة القائمة",
"Build pivots": "بناء الجداول المحورية",
"Download pivot workbook": "تنزيل مصنف الجداول المحورية",
"Pivot 1 · Category, LC/GTEE and Currency by Bucket": "الجدول المحوري 1 · الفئة وLC/GTEE والعملة حسب الفترة",
"Filters: none · Columns: Bucket · Values: Sum of Equ-IQD": "المرشحات: لا يوجد · الأعمدة: الفترة · القيم: مجموع المكافئ بالدينار العراقي",
"Pivot 2 · LC/GTEE by Currency": "الجدول المحوري 2 · LC/GTEE حسب العملة",
"Filter: CATEGORY": "المرشح: الفئة",
"Validation details": "تفاصيل التحقق",
"All": "الكل",
"Check": "الفحص",
"Value": "القيمة",
"Rows read": "الصفوف المقروءة",
"Detail rows": "صفوف التفاصيل",
"Invalid / blank Equ-IQD": "مكافئ غير صالح / فارغ",
"Source total": "إجمالي المصدر",
"Difference": "الفرق",
"Outstanding Report ready": "تقرير الأرصدة القائمة جاهز",
"Building the Outstanding Report pivots": "جارٍ بناء الجداول المحورية لتقرير الأرصدة القائمة",
"Choose the Outstanding Report workbook first": "اختر مصنف تقرير الأرصدة القائمة أولًا",
"Category": "الفئة",
"LC/GTEE": "LC/GTEE",
"Grand Total": "الإجمالي العام",
"Source workbook": "المصنف المصدر",
"Selected source sheet": "الورقة المصدر المختارة",
"Detected header row": "صف العناوين المكتشف",
"Source rows read": "صفوف المصدر المقروءة",
"Cleaned detail rows": "صفوف التفاصيل المنظّفة",
"Rows with blank Equ-IQD (excluded)": "صفوف المكافئ الفارغ (مستبعدة)",
"Rows with invalid / non-numeric Equ-IQD (excluded)": "صفوف المكافئ غير الصالح / غير الرقمي (مستبعدة)",
"Rows with blank CATEGORY": "صفوف بفئة فارغة",
"Rows with blank LC/GTEE DESC": "صفوف بوصف LC/GTEE فارغ",
"Rows with blank BILL_CCY": "صفوف بعملة الفاتورة فارغة",
"Rows with blank Bucket": "صفوف بفترة فارغة",
"Potential duplicate transaction rows (not removed)": "صفوف عمليات مكررة محتملة (لم تُحذف)",
"Source Equ-IQD total (valid rows)": "إجمالي المكافئ بالدينار في المصدر (الصفوف الصالحة)",
"Distinct CATEGORY values": "قيم الفئة المميزة",
"Distinct BILL_CCY values": "قيم عملة الفاتورة المميزة",
"Distinct Bucket values": "قيم الفترات المميزة",
"Pivot 1 Grand Total": "الإجمالي العام للجدول 1",
"Warnings": "التحذيرات",
"Search": "بحث",
"Look for every Trial Balance pivot value (BS-mapping and group level) in every non-zero cell of every sheet of each submission file - schedules and matrix sheets included. Each file gets its own workbook: your original sheets, the TB pivot, a matching report, and matched cells highlighted in the same colour on both sides.": "ابحث عن كل قيمة في الجدول المحوري لميزان المراجعة (على مستوى ربط الميزانية والمجموعة) في كل خلية غير صفرية في كل ورقة من كل ملف تقرير - بما في ذلك الجداول الملحقة وأوراق المصفوفات. يحصل كل ملف على مصنفه الخاص: أوراقك الأصلية والجدول المحوري لميزان المراجعة وتقرير المطابقة، مع تظليل الخلايا المتطابقة بنفس اللون في الجانبين.",
"Upload the Trial Balance and the submission workbooks first (the sheets do not need to be extracted).": "ارفع ميزان المراجعة ومصنفات التقارير المقدّمة أولًا (لا يلزم استخراج الأوراق).",
"1 · What is each file?": "1 · ما نوع كل ملف؟",
"Assets values (A-…) are searched in the Assets file, liabilities (L-…) in the Liabilities file, and Off-balance values (OB-… plus the Outstanding Report pivots) only in the Off-balance file. The type is guessed from the file name - change it if needed.": "تُبحث قيم الأصول (A-…) في ملف الأصول، والالتزامات (L-…) في ملف الالتزامات، وقيم خارج الميزانية (OB-… إضافةً إلى الجداول المحورية لتقرير الأرصدة القائمة) في ملف خارج الميزانية فقط. يُخمَّن النوع من اسم الملف - غيّره عند الحاجة.",
"2 · Search settings": "2 · إعدادات البحث",
"Smallest TB value to search": "أصغر قيمة من الميزان للبحث",
"Smallest TB value for a ×1,000 / ×1,000,000 match": "أصغر قيمة من الميزان لمطابقة ×1,000 / ×1,000,000",
"Tolerance (absolute)": "التفاوت (مطلق)",
"Maturity sheet keywords (Off-balance bucket pivot)": "كلمات ورقة الاستحقاق (جدول الفترات لخارج الميزانية)",
"Also try thousands / millions (sheets stated \"in thousands\")": "جرّب أيضًا الآلاف / الملايين (الأوراق المذكور فيها «بالآلاف»)",
"Accept equal-and-opposite sign": "قبول الإشارة المتساوية والمعاكسة",
"Also search individual account values (more matches, more noise)": "ابحث أيضًا في قيم الحسابات الفردية (مطابقات أكثر وضجيج أكثر)",
"Run depth search": "تشغيل البحث المتعمق",
"Download all outputs (zip)": "تنزيل كل المخرجات (zip)",
"Filter results (TB item, sheet, cell, row label, column…)": "تصفية النتائج (بند الميزان، الورقة، الخلية، تسمية الصف، العمود…)",
"Assets (033)": "الأصول (033)",
"Liabilities (034)": "الالتزامات (034)",
"Off-balance (064)": "خارج الميزانية (064)",
"Other - search everything": "أخرى - البحث في كل شيء",
"Download annotated workbook": "تنزيل المصنف المُعلَّم",
"TB item": "بند الميزان",
"TB amount": "مبلغ الميزان",
"Sheet": "الورقة",
"Cell": "الخلية",
"Sheet value": "قيمة الورقة",
"How matched": "طريقة المطابقة",
"Row label": "تسمية الصف",
"Column": "العمود",
"Currency": "العملة",
"Level": "المستوى",
"TB Amount": "مبلغ الميزان",
"exact": "مطابقة تامة",
"exact match": "مطابقة تامة",
"opposite sign": "إشارة معاكسة",
"opposite signs": "إشارتان متعاكستان",
"sign flipped": "عكس الإشارة",
"OK": "سليم",
"UNKNOWN": "غير معروف",
"CONFLICT": "تعارض",
"No matching cells.": "لا توجد خلايا مطابقة.",
"No matching cells for this filter.": "لا توجد خلايا مطابقة لهذا المرشح.",
"File type changed - run the depth search again to apply it.": "تغيّر نوع الملف - شغّل البحث المتعمق مجددًا لتطبيقه.",
"Searching every sheet for Trial Balance values": "جارٍ البحث في كل ورقة عن قيم ميزان المراجعة",
"Upload the Trial Balance first": "ارفع ميزان المراجعة أولًا",
"Upload the submission files first": "ارفع ملفات التقارير المقدّمة أولًا",
"Review the engine's suggested matches, then build or import manual matches for anything it couldn't resolve.": "راجع المطابقات التي اقترحها المحرّك، ثم أنشئ مطابقات يدوية أو استوردها لما تعذّر عليه حلّه.",
"Nothing is reconciled automatically. Every suggestion and manual match below is explicit and reviewable; nothing runs until you press Run reconciliation.": "لا شيء يُسوّى تلقائيًا. كل اقتراح ومطابقة يدوية أدناه صريح وقابل للمراجعة، ولا يُنفَّذ شيء حتى تضغط «تشغيل التسوية».",
"Show & club Trial Balance accounts by": "عرض حسابات ميزان المراجعة وتجميعها حسب",
"Account description": "وصف الحساب",
"Account no.": "رقم الحساب",
"Build a manual match": "إنشاء مطابقة يدوية",
"Pick any combination of Trial Balance items on the left and submission lines on the right - a single item, several accounts clubbed together, or a whole group against a whole sub-total. As long as the clubbed values are financially logical, the engine will sum each side and compare them.": "اختر أي تركيبة من بنود ميزان المراجعة وبنود التقارير المقدّمة - بندًا واحدًا، أو عدة حسابات مجمّعة، أو مجموعة كاملة مقابل إجمالٍ فرعي كامل. ما دامت القيم المجمّعة منطقية ماليًا فسيجمع المحرّك كل جانب ويقارنه.",
"Trial Balance items": "بنود ميزان المراجعة",
"Accounts": "الحسابات",
"BS Mapping groups (pivot level)": "مجموعات ربط الميزانية (مستوى الجدول المحوري)",
"Search account, description or BS mapping": "ابحث عن حساب أو وصف أو ربط ميزانية",
"Search BS mapping group (e.g. Assets, Other Assets)": "ابحث عن مجموعة ربط ميزانية (مثل الأصول، الأصول الأخرى)",
"Select all matching": "تحديد كل المطابق",
"Clear": "مسح",
"Selected total": "الإجمالي المحدد",
"Submission items": "بنود التقارير المقدّمة",
"Search file, sheet or line description": "ابحث عن ملف أو ورقة أو وصف بند",
"Search submissions for the TB selected total": "ابحث في التقارير عن إجمالي الميزان المحدد",
"Select Trial Balance items first, then search for that total across every extracted sheet.": "حدّد بنود ميزان المراجعة أولًا، ثم ابحث عن ذلك الإجمالي في كل ورقة مستخرجة.",
"Match description": "وصف المطابقة",
"e.g. Cash & balances (Assets) vs submission cash lines": "مثال: النقد والأرصدة (الأصول) مقابل بنود النقد في التقرير",
"Add manual match": "إضافة مطابقة يدوية",
"Manual matches": "المطابقات اليدوية",
"Save your mapping": "احفظ الربط الخاص بك",
"Download everything you've built above as a portable mapping file - re-upload it against a future period's Trial Balance and submissions to get the same matches re-resolved automatically.": "نزّل كل ما بنيته أعلاه كملف ربط قابل للنقل - أعد رفعه مع ميزان المراجعة والتقارير لفترة لاحقة لتُحَلّ المطابقات نفسها تلقائيًا.",
"Download mapping file": "تنزيل ملف الربط",
"Load a saved mapping": "تحميل ربط محفوظ",
"Upload a mapping file saved earlier. Each match is re-resolved against the current data and added below for review - nothing is applied silently.": "ارفع ملف ربط محفوظًا سابقًا. تُحَلّ كل مطابقة مجددًا مقابل البيانات الحالية وتُضاف أدناه للمراجعة - ولا يُطبَّق شيء بصمت.",
"Choose mapping file": "اختيار ملف الربط",
"Apply mapping": "تطبيق الربط",
"Absolute tolerance": "التفاوت المطلق",
"Percentage tolerance": "التفاوت النسبي",
"Run reconciliation": "تشغيل التسوية",
"Approve": "اعتماد",
"Approved": "معتمد",
"Suggestions": "الاقتراحات",
"TB groups covered": "مجموعات الميزان المغطاة",
"Avg confidence": "متوسط الثقة",
"Submission amount": "مبلغ التقرير",
"Sign applied": "الإشارة المطبّقة",
"Flipped (×-1)": "معكوسة (×-1)",
"As reported": "كما وردت",
"Auto-clubbed subset": "مجموعة فرعية مجمّعة تلقائيًا",
"File": "الملف",
"Description": "الوصف",
"Sheet amount": "مبلغ الورقة",
"Adjusted amount": "المبلغ المعدّل",
"Section match:": "مطابقة القسم:",
"Sub-group match:": "مطابقة المجموعة الفرعية:",
"Used pre-computed Total rows": "استُخدمت صفوف الإجمالي الجاهزة",
"Summed individual lines": "جُمعت البنود الفردية",
"No candidate matches were generated. Build a manual match below.": "لم تُولَّد مطابقات مرشحة. أنشئ مطابقة يدوية أدناه.",
"No manual matches yet - build one above or load a saved mapping file.": "لا توجد مطابقات يدوية بعد - أنشئ واحدة أعلاه أو حمّل ملف ربط محفوظًا.",
"Remove": "إزالة",
"+ Add": "+ إضافة",
"Added ✓": "تمت الإضافة ✓",
"Total:": "الإجمالي:",
"none resolved": "لم يُحَل أي شيء",
"Forex": "العملات الأجنبية",
"BS Mapping group total (pivot level)": "إجمالي مجموعة ربط الميزانية (مستوى الجدول المحوري)",
"No BS Mapping groups yet - upload a Trial Balance first.": "لا توجد مجموعات ربط ميزانية بعد - ارفع ميزان المراجعة أولًا.",
"No Trial Balance items yet - upload a Trial Balance first.": "لا توجد بنود ميزان مراجعة بعد - ارفع ميزان المراجعة أولًا.",
"No submission items yet - extract sheets first.": "لا توجد بنود تقارير بعد - استخرج الأوراق أولًا.",
"Sign flipped": "إشارة معكوسة",
"Submissions extracted": "تم استخراج التقارير",
"Extracting submission sheets": "جارٍ استخراج أوراق التقارير",
"Select at least one sheet": "حدّد ورقة واحدة على الأقل",
"Of every mapping rule the app knows about (the bundled default mapping, plus anything you've imported), how many actually found a match - and how many didn't, and why.": "من بين كل قاعدة ربط يعرفها التطبيق (الربط الافتراضي المضمّن إضافةً إلى ما استوردته)، كم منها وجد مطابقة وكم منها لم يجد، ولماذا.",
"Export workbook": "تصدير المصنف",
"Every rule, in detail": "كل قاعدة بالتفصيل",
"Fulfilled = data was found on both sides. Reconciled = it was also approved and run, and the numbers agreed.": "مُنجَزة = وُجدت البيانات في الجانبين. مسوّاة = اعتُمدت أيضًا ونُفّذت وتطابقت الأرقام.",
"Fulfilled only": "المنجزة فقط",
"Unresolved only": "غير المحلولة فقط",
"Reconciled only": "المسوّاة فقط",
"Needs review only": "المحتاجة لمراجعة فقط",
"Search label, currency or warning text": "ابحث في التسمية أو العملة أو نص التحذير",
"Total rules": "إجمالي القواعد",
"Fulfilled (data found)": "منجزة (وُجدت البيانات)",
"Unresolved (no data)": "غير محلولة (لا بيانات)",
"Reconciled (matched)": "مسوّاة (متطابقة)",
"Needs review": "تحتاج مراجعة",
"Mapping rule": "قاعدة الربط",
"Data found?": "هل وُجدت البيانات؟",
"TB accounts": "حسابات الميزان",
"Submission lines": "بنود التقارير",
"Reconciliation outcome": "نتيجة التسوية",
"Submission Amount": "مبلغ التقرير",
"Detail": "التفاصيل",
"FULFILLED": "منجزة",
"NOT FOUND": "غير موجودة",
"not run": "لم تُنفَّذ",
"Reconciliation Results": "نتائج التسوية",
"Traceable amounts, variance and review status per group, sub-group and currency.": "مبالغ قابلة للتتبع وفروقات وحالة مراجعة لكل مجموعة ومجموعة فرعية وعملة.",
"How every Trial Balance group connects to its submission evidence - the full lineage, end to end.": "كيف ترتبط كل مجموعة في ميزان المراجعة بأدلة التقرير المقدّم - التتبع الكامل من البداية إلى النهاية.",
"Audit-ready workbook": "مصنف جاهز للتدقيق",
"Includes executive summary, TB pivot, TB detail, submission inventory, auto-suggestions, reconciliation results and the full reconciliation map.": "يتضمن الملخص التنفيذي والجدول المحوري لميزان المراجعة وتفاصيله وجرد التقارير المقدّمة والاقتراحات التلقائية ونتائج التسوية وخريطة التسوية الكاملة.",
"Download reconciliation workbook": "تنزيل مصنف التسوية",
"Temporary session files are not written to a Dataiku managed folder.": "لا تُكتب ملفات الجلسة المؤقتة في مجلد Dataiku المُدار.",
"BS Mapping": "ربط الميزانية",
"Variance %": "نسبة الفرق %",
"Status": "الحالة",
"Rule": "القاعدة",
"Source": "المصدر",
"Account": "الحساب",
"Row": "الصف",
"Amount": "المبلغ",
"Section": "القسم",
"Line description": "وصف البند",
"Total row?": "صف إجمالي؟",
"Trial Balance side": "جانب ميزان المراجعة",
"Submission side": "جانب التقرير المقدّم",
"No evidence rows": "لا توجد صفوف أدلة",
"No reconciliation rule approved for this line yet.": "لم تُعتمد أي قاعدة تسوية لهذا البند بعد.",
"Manual & imported matches": "المطابقات اليدوية والمستوردة",
"Run the reconciliation to build the map.": "شغّل التسوية لبناء الخريطة.",
"Reconciliation complete": "اكتملت التسوية",
"Approve at least one suggestion or build a manual match": "اعتمد اقتراحًا واحدًا على الأقل أو أنشئ مطابقة يدوية",
"No records yet": "لا توجد سجلات بعد",
"MATCH": "مطابق",
"MATCH_WITHIN_TOLERANCE": "مطابق ضمن التفاوت",
"REVIEW_REQUIRED": "يتطلب مراجعة",
"UNRESOLVED": "غير محلول",
"NOT_APPROVED": "غير معتمد",
"MANUAL": "يدوي",
"AUTO": "تلقائي",
"DEFAULT": "افتراضي",
"IMPORTED": "مستورد",
"TOTAL": "الإجمالي",
"Session ready": "الجلسة جاهزة",
"Could not start a session": "تعذّر بدء الجلسة",
"Error - see message": "خطأ - راجع الرسالة",
"Processing Trial Balance": "جارٍ معالجة ميزان المراجعة",
"TB pivot ready": "الجدول المحوري للميزان جاهز",
"Choose a Trial Balance workbook first": "اختر مصنف ميزان المراجعة أولًا",
"Choose submission files first": "اختر ملفات التقارير المقدّمة أولًا",
"Inspecting submission workbooks": "جارٍ فحص مصنفات التقارير",
"Choose submission sheets": "اختر أوراق التقارير",
"Choose a working language": "اختر لغة العمل",
"Choose Arabic or English first": "اختر العربية أو الإنجليزية أولًا",
"Choose your dictionary workbook first": "اختر مصنف القاموس أولًا",
"Translating the submissions to English": "جارٍ ترجمة التقارير إلى الإنجليزية",
"No active session yet": "لا توجد جلسة نشطة بعد",
"Select Trial Balance items first": "حدّد بنود ميزان المراجعة أولًا",
"Extract submission sheets first": "استخرج أوراق التقارير أولًا",
"Added to submission side of the match": "أُضيف إلى جانب التقرير في المطابقة",
"Select at least one item on each side": "حدّد بندًا واحدًا على الأقل في كل جانب",
"Give this match a short description": "أعطِ هذه المطابقة وصفًا قصيرًا",
"Selected TB items span multiple currencies - compared as a combined TOTAL.": "بنود الميزان المحددة تشمل عدة عملات - ستُقارَن كإجمالٍ مجمّع.",
"Manual match added": "تمت إضافة المطابقة اليدوية",
"No manual matches to save yet": "لا توجد مطابقات يدوية للحفظ بعد",
"Choose a mapping file first": "اختر ملف الربط أولًا",
"That file is not valid JSON.": "هذا الملف ليس JSON صالحًا.",
"No matches found in that mapping file.": "لم يُعثر على مطابقات في ملف الربط هذا.",
"Some picks only partly covered the new grouping and were dropped - re-pick them.": "بعض الاختيارات غطّت التجميع الجديد جزئيًا فأُسقطت - أعد اختيارها.",
"The bahrain_iraq_algorithm library in your project is out of date (no FRX* support) - copy the latest engine.py into it and restart the backend.": "مكتبة bahrain_iraq_algorithm في مشروعك قديمة (لا تدعم FRX*) - انسخ أحدث engine.py إليها وأعد تشغيل الخلفية.",
"This usually means the server raised an error. Check the Dataiku backend log.": "يعني هذا غالبًا أن الخادم أصدر خطأً. راجع سجل خلفية Dataiku.",
"Upload a Trial Balance workbook.": "ارفع مصنف ميزان المراجعة.",
"Upload at least one submission workbook.": "ارفع مصنف تقارير مقدّمة واحدًا على الأقل.",
"Upload submission workbooks first.": "ارفع مصنفات التقارير المقدّمة أولًا.",
"Upload the submission workbooks first.": "ارفع مصنفات التقارير المقدّمة أولًا.",
"Upload the submission workbooks first (Submissions step).": "ارفع مصنفات التقارير المقدّمة أولًا (خطوة التقارير المقدّمة).",
"Upload and process the Trial Balance first.": "ارفع ميزان المراجعة وعالجه أولًا.",
"Process the Trial Balance and submissions first.": "عالج ميزان المراجعة والتقارير المقدّمة أولًا.",
"Process submissions first.": "عالج التقارير المقدّمة أولًا.",
"No processed Trial Balance.": "لا يوجد ميزان مراجعة معالَج.",
"No submission lines extracted yet.": "لم تُستخرج بنود تقارير بعد.",
"No amount given.": "لم يُحدَّد مبلغ.",
"No file uploaded": "لم يُرفع أي ملف",
"Run reconciliation first.": "شغّل التسوية أولًا.",
"Run the depth search first.": "شغّل البحث المتعمق أولًا.",
"That file was not part of the last depth search.": "لم يكن هذا الملف ضمن آخر بحث متعمق.",
"No searched files to download.": "لا توجد ملفات مبحوثة للتنزيل.",
"Upload the Outstanding Report first.": "ارفع تقرير الأرصدة القائمة أولًا.",
"Upload the Outstanding Report workbook.": "ارفع مصنف تقرير الأرصدة القائمة.",
"Choose Arabic or English.": "اختر العربية أو الإنجليزية.",
"No translation has been run yet.": "لم تُشغَّل أي ترجمة بعد.",
"That file was not translated.": "لم يُترجَم هذا الملف.",
"Upload your dictionary workbook (Arabic in column A, English in column B, header in row 1).": "ارفع مصنف القاموس (العربية في العمود A والإنجليزية في العمود B والعناوين في الصف 1).",
"The dictionary has no usable rows (Arabic in column A, English in column B, header in row 1).": "لا يحتوي القاموس على صفوف صالحة (العربية في العمود A والإنجليزية في العمود B والعناوين في الصف 1).",
"Upload a Trial Balance and process submissions before importing a mapping.": "ارفع ميزان المراجعة وعالج التقارير قبل استيراد ربط.",
"The uploaded mapping file has no matches in it.": "لا يحتوي ملف الربط المرفوع على أي مطابقات.",
"Repeated header rows removed": "صفوف العناوين المكررة المحذوفة",
"Bucket labels that differ only by case/spacing": "تسميات الفترات التي تختلف في الحروف/المسافات فقط",
"Sum of Equ-IQD by Bucket": "مجموع المكافئ بالدينار حسب الفترة",
"Pivot 2 Grand Total (all categories)": "الإجمالي العام للجدول 2 (كل الفئات)",
"Pivot 2 Grand Total (as selected)": "الإجمالي العام للجدول 2 (حسب الاختيار)",
"Reconciliation difference (source - Pivot 1)": "فرق التسوية (المصدر - الجدول 1)",
"Reconciliation difference (source - Pivot 2)": "فرق التسوية (المصدر - الجدول 2)",
"· Sub-group match:": "· مطابقة المجموعة الفرعية:",
"· Summed individual lines": "· جُمعت البنود الفردية",
"· Used pre-computed Total rows": "· استُخدمت صفوف الإجمالي الجاهزة",
"Forex: sub-total of every currency other than the local currency": "العملات الأجنبية: المجموع الفرعي لكل عملة غير العملة المحلية",
"Amount search": "البحث بالمبلغ",
"Sheets": "الأوراق",
"Rows with invalid or blank Equ-IQD": "صفوف بمكافئ غير صالح أو فارغ",
"Distinct values": "القيم المميزة",
"Step 3 · Off-balance": "الخطوة 3 · خارج الميزانية",
"Off-balance pivots (optional)": "جداول خارج الميزانية المحورية (اختياري)",
"Optional: upload it after the Trial Balance and before the submissions. It is used only to reconcile the Off-balance (064) submission. The raw-data sheet and header row are found automatically, Equ-IQD is cleaned and validated, and two pivots are built. Depth Search then looks for their values in the Off-balance submission only, with the Bucket pivot searched on its Maturity sheet only.": "اختياري: ارفعه بعد ميزان المراجعة وقبل التقارير المقدّمة. يُستخدم فقط لتسوية تقرير خارج الميزانية (064). تُكتشف ورقة البيانات الخام وصف العناوين تلقائيًا، وتُنظَّف قيمة المكافئ بالدينار العراقي وتُراجع، ويُبنى جدولان محوريان. ثم يبحث البحث المتعمق عن قيمهما في تقرير خارج الميزانية فقط، ويُبحث في الجدول المحوري للفترات في ورقة الاستحقاق فقط.",
"Default Mapping": "الربط الافتراضي",
"Rules": "القواعد",
"The bundled mapping is applied to the Trial Balance and every submission file as soon as the submissions are extracted. Each file gets its own workbook — your original sheets, the TB pivot and accounts, a report that states each rule, the Trial Balance and submission amounts and the variance, and the matched cells highlighted in one colour on both sides. The same results can be downloaded as a PDF.": "يُطبَّق الربط المدمج على ميزان المراجعة وعلى كل ملف مقدَّم فور استخراج التقارير. يحصل كل ملف على مصنفه الخاص: أوراقك الأصلية، والجدول المحوري للحسابات، وتقرير يوضح كل قاعدة ومبالغ ميزان المراجعة والتقرير المقدَّم والفرق، مع تمييز الخلايا المطابقة بلون واحد على الجانبين. ويمكن تنزيل النتائج نفسها بصيغة PDF.",
"Upload the Trial Balance, then upload and extract the submission workbooks (Submissions step). The default mapping runs automatically.": "ارفع ميزان المراجعة، ثم ارفع مصنفات التقارير المقدَّمة واستخرجها (خطوة التقارير المقدَّمة). يعمل الربط الافتراضي تلقائيًا.",
"Tolerance (% of the TB amount)": "التفاوت (% من مبلغ ميزان المراجعة)",
"Re-run with these tolerances": "إعادة التشغيل بهذه التفاوتات",
"Download PDF (all results)": "تنزيل PDF (كل النتائج)",
"Download all workbooks (zip)": "تنزيل كل المصنفات (zip)",
"Filter rules (name, how it works, sheet, cell…)": "تصفية القواعد (الاسم، طريقة العمل، الورقة، الخلية…)",
"All statuses": "كل الحالات",
"Variance - review": "فرق - يتطلب مراجعة",
"Match (within tolerance)": "مطابق (ضمن التفاوت)",
"Match": "مطابق",
"Not found": "غير موجود",
"Rules applied": "القواعد المطبقة",
"How the rule works": "طريقة عمل القاعدة",
"Variance": "الفرق",
"Where in the file": "الموضع في الملف",
"(not found)": "(غير موجود)",
"Rules that apply to none of the uploaded files": "قواعد لا تنطبق على أي من الملفات المرفوعة",
"No rules for this filter.": "لا توجد قواعد لهذه التصفية.",
"No rules to show.": "لا توجد قواعد للعرض.",
"Variance (TB - submission):": "الفرق (ميزان المراجعة - التقرير المقدَّم):",
"Re-running the default mapping": "جارٍ إعادة تشغيل الربط الافتراضي",
"Process the submissions first.": "عالج التقارير المقدَّمة أولًا.",
"That file is not part of this session.": "هذا الملف ليس ضمن هذه الجلسة.",
"No submission file has a default-mapping result.": "لا يوجد ملف مقدَّم له نتيجة ربط افتراضي.",
"Trial Balance groups and the rules that read them": "مجموعات ميزان المراجعة والقواعد التي تقرأها",
"BS mapping group": "مجموعة ربط الميزانية",
"Foreign currencies": "العملات الأجنبية",
"Rules that read it": "القواعد التي تقرأها",
"NO RULE": "لا توجد قاعدة"
};
const AR_RULES=[
[
"^(\\d+) selected$",
"{1} محدد"
],
[
"^(\\d+) selected \\((\\d+) accounts\\)$",
"{1} محدد ({2} حسابات)"
],
[
"^Showing first (\\d+) of (\\d+) - refine your search$",
"عرض أول {1} من {2} - حسّن بحثك"
],
[
"^Showing (\\d+) of (\\d+) - download the workbook for the full report\\.$",
"عرض {1} من {2} - نزّل المصنف للحصول على التقرير الكامل."
],
[
"^(\\d+) sheet\\(s\\)$",
"{1} ورقة"
],
[
"^(\\d+) sheets scanned$",
"تم فحص {1} ورقة"
],
[
"^([\\d,.]+) non-zero cells$",
"{1} خلية غير صفرية"
],
[
"^(\\d+) of (\\d+) TB values found$",
"تم العثور على {1} من {2} قيمة من الميزان"
],
[
"^(\\d+) currency conflict\\(s\\) not highlighted$",
"{1} تعارض في العملة غير مظلَّل"
],
[
"^(\\d+) TB value\\(s\\) not found in this file$",
"{1} قيمة من الميزان لم يُعثر عليها في هذا الملف"
],
[
"^Depth search done - (\\d+) TB value\\(s\\) located$",
"اكتمل البحث المتعمق - تم تحديد {1} قيمة"
],
[
"^(\\d+) accts?$",
"{1} حساب"
],
[
"^(\\d+) lines · (\\d+) accts$",
"{1} بنود · {2} حسابات"
],
[
"^(\\d+) lines$",
"{1} بنود"
],
[
"^(\\d+) lines · (\\d+) accts$",
"{1} بنود · {2} حسابات"
],
[
"^confidence (\\d+)%$",
"الثقة {1}%"
],
[
"^Match basis & evidence \\((\\d+) lines?\\)$",
"أساس المطابقة والأدلة ({1} بنود)"
],
[
"^(\\d+) match(?:es)?$",
"{1} مطابقة"
],
[
"^(\\d+) TB account row\\(s\\)$",
"{1} صف حساب من الميزان"
],
[
"^(\\d+) account rows$",
"{1} صفوف حسابات"
],
[
"^Inspected (\\d+) workbook\\(s\\)\\. Choose the sheets to read, then extract\\.$",
"تم فحص {1} مصنف. اختر الأوراق المراد قراءتها ثم نفّذ الاستخراج."
],
[
"^Extracted$",
"تم استخراج"
],
[
"^currency-tagged lines\\. Generated$",
"بنود موسومة بالعملة. تم توليد"
],
[
"^candidate matches\\.$",
"مطابقات مرشحة."
],
[
"^candidate matches\\. Included$",
"مطابقات مرشحة. تم تضمين"
],
[
"^match(?:es)? from the default mapping automatically \\(review under Manual matches - see Mapping Coverage for the full picture: (.*)\\)\\.$",
"مطابقة من الربط الافتراضي تلقائيًا (راجعها ضمن المطابقات اليدوية - وانظر تغطية الربط للصورة الكاملة: {1:t})."
],
[
"^(\\d+) of (\\d+) rules fulfilled$",
"{1} من {2} قاعدة منجزة"
],
[
"^, (\\d+) data rows\\.$",
"، {1} صفوف بيانات."
],
[
"^Translated (\\d+) cell\\(s\\) and (\\d+) sheet name\\(s\\) using (\\d+) dictionary entries\\. (\\d+) Arabic term\\(s\\) were not in the dictionary - the English workbook is used from here on\\.$",
"تمت ترجمة {1} خلية و{2} اسم ورقة باستخدام {3} مدخلًا من القاموس. {4} مصطلحًا عربيًا لم يكن في القاموس - يُستخدم المصنف الإنجليزي من الآن فصاعدًا."
],
[
"^- (\\d+) Arabic cell\\(s\\) in (\\d+) sheet\\(s\\)$",
"- {1} خلية عربية في {2} ورقة"
],
[
"^Will search for: (.+) \\(current TB selected total\\)$",
"سيتم البحث عن: {1} (إجمالي الميزان المحدد حاليًا)"
],
[
"^Found (\\d+) line\\(s\\) across all extracted sheets matching (.+)\\.$",
"تم العثور على {1} بند في كل الأوراق المستخرجة يطابق {2}."
],
[
"^Found (\\d+) line\\(s\\) across all extracted sheets matching (.+) \\(showing the closest 200\\)\\.$",
"تم العثور على {1} بند في كل الأوراق المستخرجة يطابق {2} (عرض أقرب 200)."
],
[
"^Found (\\d+) match\\(es\\)$",
"تم العثور على {1} مطابقة"
],
[
"^Detected (.+) - applied automatically\\.$",
"تم اكتشاف {1:t} - طُبّق تلقائيًا."
],
[
"^([×÷][\\d,.]+) scale$",
"مقياس {1}"
],
[
"^([×÷][\\d,.]+) scale and sign flipped$",
"مقياس {1} مع عكس الإشارة"
],
[
"^a ([×÷][\\d,.]+) scale$",
"مقياس {1}"
],
[
"^a ([×÷][\\d,.]+) scale and sign flipped$",
"مقياس {1} مع عكس الإشارة"
],
[
"^✓ Values match directly - TB (.+) vs submission (.+)\\.$",
"✓ القيم متطابقة مباشرة - الميزان {1} مقابل التقرير {2}."
],
[
"^✓ Equal and opposite - TB (.+) vs submission (.+)\\. The submission side will be sign-flipped when you add the match\\.$",
"✓ متساويان ومتعاكسان - الميزان {1} مقابل التقرير {2}. ستُعكس إشارة جانب التقرير عند إضافة المطابقة."
],
[
"^⚠ TB (.+) vs submission (.+) don't match directly, but with the (.+) the submission becomes (.+) - this will be applied automatically when you add the match\\.$",
"⚠ الميزان {1} مقابل التقرير {2} غير متطابقين مباشرة، لكن مع {3:t} يصبح التقرير {4} - وسيُطبَّق ذلك تلقائيًا عند إضافة المطابقة."
],
[
"^✓ Values match - TB (.+) vs submission (.+) \\(using the scale/sign found by amount search\\)\\.$",
"✓ القيم متطابقة - الميزان {1} مقابل التقرير {2} (باستخدام المقياس/الإشارة التي وجدها البحث بالمبلغ)."
],
[
"^⚠ TB (.+) vs submission (.+) \\(after the amount-search transform\\) still don't agree - double-check before adding\\.$",
"⚠ الميزان {1} مقابل التقرير {2} (بعد تحويل البحث بالمبلغ) لا يزالان غير متطابقين - تحقق مجددًا قبل الإضافة."
],
[
"^Imported (\\d+) match\\(es\\), (\\d+) need review \\(see the warning badges below\\)\\.$",
"تم استيراد {1} مطابقة، منها {2} تحتاج مراجعة (انظر شارات التحذير أدناه)."
],
[
"^Imported (\\d+) match\\(es\\)\\.$",
"تم استيراد {1} مطابقة."
],
[
"^Trial Balance \\((\\d+)\\)$",
"ميزان المراجعة ({1})"
],
[
"^Trial Balance \\((\\d+) by (description|account no\\.) · (\\d+) lines\\)$",
"ميزان المراجعة ({1} حسب {2:t} · {3} بنود)"
],
[
"^description$",
"الوصف"
],
[
"^Submission \\((\\d+)\\)$",
"التقرير المقدّم ({1})"
],
[
"^Scale (.+) applied$",
"تم تطبيق المقياس {1}"
],
[
"^Scale-adjusted (.+)$",
"مُعدَّل بالمقياس {1}"
],
[
"^TOTAL ([\\d,.\\-]+)$",
"الإجمالي {1}"
],
[
"^All non-(\\w+) currencies, re-totalled each time the mapping is applied$",
"كل العملات غير {1}، يُعاد جمعها في كل مرة يُطبَّق فيها الربط"
],
[
"^Download (.+) \\(translated\\)$",
"تنزيل {1} (مترجم)"
],
[
"^([×÷][\\d,.]+) \\(sheet is in thousands\\)$",
"{1} (الورقة بالآلاف)"
],
[
"^This column is a different currency from the TB value$",
"هذا العمود بعملة مختلفة عن قيمة الميزان"
],
[
"^(\\d+) sheets/files had this section$",
"{1} أوراق/ملفات تحتوي هذا القسم"
],
[
"^- picked the one whose amount agrees with the TB: (.+?) · (.+)$",
"- تم اختيار الذي يتفق مبلغه مع الميزان: {1} · {2:t}"
],
[
"^- none agreed with the TB, took the closest: (.+?) · (.+)$",
"- لم يتفق أي منها مع الميزان، وتم أخذ الأقرب: {1} · {2:t}"
],
[
"^· Auto-clubbed a subset of lines \\(excluded (\\d+)\\)$",
"· جُمّعت مجموعة فرعية من البنود (استُبعد {1})"
],
[
"^· Scale detected: (.+)$",
"· المقياس المكتشف: {1}"
],
[
"^Auto-clubbed a subset of lines \\(excluded (\\d+)\\)$",
"جُمّعت مجموعة فرعية من البنود (استُبعد {1})"
],
[
"^TB (.+) vs submission (.+) \\(Δ (.+)\\)$",
"الميزان {1} مقابل التقرير {2} (Δ {3})"
],
[
"^· OB Pivot (\\d)$",
"· جدول خارج الميزانية {1}"
],
[
"^No sheet in this file has a maturity/tenure name \\((.*)\\), so the Outstanding Report bucket pivot \\((\\d+) values\\) was not searched here\\. Adjust the keywords if the sheet is named differently\\.$",
"لا توجد ورقة في هذا الملف باسم استحقاق/أجل ({1})، لذلك لم يُبحث هنا عن الجدول المحوري للفترات في تقرير الأرصدة القائمة ({2} قيمة). عدّل الكلمات إذا كانت الورقة تحمل اسمًا مختلفًا."
],
[
"^(\\d+) formula cell\\(s\\) have no saved value \\(the file was never recalculated in Excel\\), so they could not be searched - open and save the file in Excel once, then upload it again\\.$",
"{1} خلية معادلة بلا قيمة محفوظة (لم يُعَد حساب الملف في Excel قط) لذلك تعذّر البحث فيها - افتح الملف في Excel واحفظه مرة واحدة ثم ارفعه مجددًا."
],
[
"^Could not read this workbook: (.+)$",
"تعذّرت قراءة هذا المصنف: {1}"
],
[
"^Off-balance files are not searched yet\\.$",
"ملفات خارج الميزانية لا تُبحث بعد."
],
[
"^Outstanding Report loaded \\((.+)\\): it is searched in the Off-balance \\(064\\) file only - its Pivot 1 \\(Bucket\\) on Maturity sheets, and Pivot 2 \\(CATEGORY: (.+)\\) on every sheet\\.$",
"تم تحميل تقرير الأرصدة القائمة ({1}): يُبحث عنه في ملف خارج الميزانية (064) فقط - الجدول المحوري 1 (الفترات) في أوراق الاستحقاق، والجدول المحوري 2 (الفئة: {2:t}) في كل ورقة."
],
[
"^Outstanding Report loaded \\((.+)\\), but no file is typed Off-balance \\(064\\) - set the type of the Off-balance submission below, otherwise its pivots are not searched\\.$",
"تم تحميل تقرير الأرصدة القائمة ({1})، لكن لا يوجد ملف مصنّف كخارج الميزانية (064) - حدّد نوع تقرير خارج الميزانية أدناه، وإلا فلن تُبحث جداوله المحورية."
],
[
"^The Outstanding Report pivots were not searched: no file is typed Off-balance \\(064\\)\\. Set the type of the Off-balance submission under \\\"What is each file\\?\\\" and run again\\.$",
"لم تُبحث الجداول المحورية لتقرير الأرصدة القائمة: لا يوجد ملف مصنّف كخارج الميزانية (064). حدّد نوع تقرير خارج الميزانية تحت «ما نوع كل ملف؟» ثم أعد التشغيل."
],
[
"^Searched: (\\d+) OB- Trial Balance values \\+ (\\d+) Outstanding Report values$",
"تم البحث عن: {1} قيمة OB- من ميزان المراجعة + {2} قيمة من تقرير الأرصدة القائمة"
],
[
"^(\\d+) rules$",
"{1} قاعدة"
],
[
"^(\\d+) matched$",
"{1} مطابقة"
],
[
"^(\\d+) with a variance to review$",
"{1} بفرق يتطلب مراجعة"
],
[
"^(\\d+) not found$",
"{1} غير موجودة"
],
[
"^Total absolute variance ([\\d,.\\-]+)$",
"إجمالي الفرق المطلق {1}"
],
[
"^Total absolute variance on rules that matched with a difference: (.+)\\. Variance = Trial Balance amount - submission amount\\.$",
"إجمالي الفرق المطلق للقواعد التي طابقت بوجود فرق: {1}. الفرق = مبلغ ميزان المراجعة - مبلغ التقرير المقدَّم."
],
[
"^Default mapping: (\\d+) variance\\(s\\) to review$",
"الربط الافتراضي: {1} فرق (فروق) تتطلب مراجعة"
],
[
"^(\\d+) of (\\d+) groups covered$",
"{1} من {2} مجموعة مغطاة"
],
[
"^(\\d+) group\\(s\\) with no rule$",
"{1} مجموعة بلا قاعدة"
],
[
"^The sheet '(.+)' was not extracted, so the checks that read it were skipped\\.$",
"لم تُستخرج الورقة '{1}'، لذلك تم تخطي الفحوصات التي تقرأها."
]
].map(([src,tpl])=>[new RegExp(src),tpl]);
const I18N_ATTRS=['placeholder','title','aria-label'];
const I18N_SKIP_SEL='td:not([data-tr]),.pTitle,.pSub,.nodeName,.acctRow,.srTitle,.srSub,.sheetTitle,.fileCardHead b,.depthFileRow b,.chipFile';
const I18N_TEXT=new WeakMap();      // text node -> {en, arRaw}
const I18N_ATTR=new WeakMap();      // element   -> {attr: {en, ar}}
let i18nObs=null, i18nBusy=false;

function i18nKey(s){ return String(s).replace(/[‐-―−]/g,'-').replace(/ /g,' ').replace(/\s+/g,' ').trim(); }
const AR_NORM=Object.create(null);
Object.keys(AR_DICT).forEach(k=>{ AR_NORM[i18nKey(k)]=AR_DICT[k]; });

function translateUi(s,depth){
  const k=i18nKey(s); if(!k) return null;
  if(AR_NORM[k]!==undefined) return AR_NORM[k];
  for(const [re,tpl] of AR_RULES){
    const m=re.exec(k);
    if(m) return tpl.replace(/\{(\d)(:t)?\}/g,(_,i,t)=>{
      const g=m[+i]||''; if(!t) return g;
      const x=translateUi(g,(depth||0)+1); return x==null?g:x; });
  }
  if(!depth && k.indexOf(', ')>0){                      // "x1,000 (sheet is in thousands), opposite sign"
    const parts=k.split(', '), tr=parts.map(p=>translateUi(p,1));
    if(tr.some(x=>x!=null)) return parts.map((p,i)=>tr[i]==null?p:tr[i]).join('، ');
  }
  return null;
}
function i18nSkippable(el){
  if(!el) return true;
  if(el.closest('script,style,.langBtn')) return true;
  if(el.closest('.pill,.chip,[data-tr]')) return false;
  return !!el.closest(I18N_SKIP_SEL);
}
function i18nText(n){
  const raw=n.nodeValue; if(!raw||!raw.trim()) return;
  const rec=I18N_TEXT.get(n);
  if(state.uiLang==='ar'){
    if(rec && raw===rec.arRaw) return;                 // already translated by us
    if(i18nSkippable(n.parentElement)) return;
    const ar=translateUi(raw); if(ar==null) return;
    const out=raw.match(/^\s*/)[0]+ar+raw.match(/\s*$/)[0];
    I18N_TEXT.set(n,{en:raw,arRaw:out}); n.nodeValue=out;
  }else if(rec){
    if(raw===rec.arRaw) n.nodeValue=rec.en;
    I18N_TEXT.delete(n);
  }
}
function i18nAttrs(el){
  if(!el.getAttribute) return;
  for(const a of I18N_ATTRS){
    const v=el.getAttribute(a); if(!v) continue;
    const m=I18N_ATTR.get(el)||{}, rec=m[a];
    if(state.uiLang==='ar'){
      if(rec && v===rec.ar) continue;
      const ar=translateUi(v); if(ar==null) continue;
      m[a]={en:v,ar}; I18N_ATTR.set(el,m); el.setAttribute(a,ar);
    }else if(rec){
      if(v===rec.ar) el.setAttribute(a,rec.en);
      delete m[a];
    }
  }
}
function i18nWalk(root){
  if(root.nodeType===1){ i18nAttrs(root); root.querySelectorAll('[placeholder],[title],[aria-label]').forEach(i18nAttrs); }
  const tw=document.createTreeWalker(root,NodeFilter.SHOW_TEXT,null);
  const nodes=[]; while(tw.nextNode()) nodes.push(tw.currentNode);
  nodes.forEach(i18nText);
}
function i18nStartObserver(){
  if(i18nObs||!window.MutationObserver) return;
  i18nObs=new MutationObserver(muts=>{
    if(i18nBusy||state.uiLang!=='ar') return;
    i18nBusy=true;
    try{
      muts.forEach(m=>{
        if(m.type==='characterData') i18nText(m.target);
        else if(m.type==='attributes') i18nAttrs(m.target);
        else m.addedNodes.forEach(nd=>{ if(nd.nodeType===3) i18nText(nd); else if(nd.nodeType===1) i18nWalk(nd); });
      });
    }finally{ i18nBusy=false; i18nObs.takeRecords(); }
  });
  i18nObs.observe(document.body,{childList:true,subtree:true,characterData:true,attributes:true,attributeFilter:I18N_ATTRS});
}
function applyLanguage(){
  const ar=state.uiLang==='ar';
  const app=document.getElementById('app');
  if(app){ app.setAttribute('dir',ar?'rtl':'ltr'); app.setAttribute('lang',ar?'ar':'en'); }
  document.querySelectorAll('#langSwitch .langBtn').forEach(b=>b.classList.toggle('active',b.dataset.lang===state.uiLang));
  i18nBusy=true;
  try{ i18nWalk(document.body); }finally{ i18nBusy=false; if(i18nObs) i18nObs.takeRecords(); }
  i18nStartObserver();
}
function loadSavedLang(){
  try{ if(localStorage.getItem('recon_ui_lang')==='ar') state.uiLang='ar'; }catch(e){}
  applyLanguage();
}

loadSavedLang();
init();
