/* Bahrain-Iraq Recon Studio — frontend.
   No inline event handlers anywhere (CSP-safe): every interaction is wired
   through addEventListener + event delegation, and every network call goes
   through api() which never lets a non-JSON / error response fail silently. */

let state={session:null,pivot:[],pivotCols:[],tbTree:[],tbItems:[],tbSel:new Set(),
  tbGroupItems:[],tbSelGroups:new Set(),tbLevel:'account',
  subFilesMeta:[],selection:{},subPreview:[],subSel:new Set(),
  suggestions:[],suggestionDecisions:{},manualMatches:[],results:[],lineage:[]};

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
  let original;
  if(busyBtn){ original=busyBtn.innerHTML; busyBtn.disabled=true; busyBtn.classList.add('isBusy'); }
  try{
    await fn();
  }catch(err){
    console.error(err);
    toast((err&&err.message)||String(err));
    setStatus('Error — see message',false);
  }finally{
    if(busyBtn){ busyBtn.disabled=false; busyBtn.classList.remove('isBusy'); if(original!=null) busyBtn.innerHTML=original; }
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
  state.tbItems=flattenTbItems(state.tbTree); state.tbSel=new Set();
  state.tbGroupItems=flattenTbGroups(state.tbTree); state.tbSelGroups=new Set();
  $('tbMessage').innerHTML=`<div class="notice good">Detected sheet <b>${esc(j.meta.sheet)}</b>, header row <b>${j.meta.header_row}</b>, ${j.meta.rows} data rows.</div>`;
  $('kpis').innerHTML=Object.entries(j.kpis).map(([k,v])=>`<div class="kpi"><small>${esc(k.replace('_',' '))}</small><b>${esc(fmt(v))}</b></div>`).join('');
  renderPivot(); renderTbTree(); renderTbPicker(); setStatus('TB pivot ready'); go('pivot');
}
function renderPivot(){
  const q=($('pivotSearch')?.value||'').toLowerCase();
  const rows=(state.pivot||[]).filter(r=>JSON.stringify(r).toLowerCase().includes(q));
  table('pivotTable',rows,(state.pivotCols||[]).map(c=>({key:c,label:c})));
}
function currencyChips(totals){
  if(!totals) return '';
  return Object.entries(totals).filter(([k])=>k!=='TOTAL').map(([k,v])=>`<span class="chip">${esc(k)} ${esc(fmt(v))}</span>`).join('')+
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

async function uploadSubs(){
  const fs=[...$('subFiles').files]; if(!fs.length){toast('Choose submission files first');return}
  const fd=new FormData(); fd.append('session_id',state.session); fs.slice(0,3).forEach(f=>fd.append('files',f));
  setStatus('Inspecting submission workbooks');
  const j=await api('/api/submissions/upload',{method:'POST',body:fd});
  state.subFilesMeta=j.files||[]; state.selection={};
  state.subFilesMeta.forEach(f=>{ state.selection[f.file]={};
    (f.sheets||[]).forEach(s=>{ state.selection[f.file][s.sheet]={checked: s.header_row!=null && s.score>=3, header_row: s.header_row}; }); });
  renderSheetPicker();
  $('extractActions').style.display='flex';
  $('subDownloadRow').style.display='none';
  $('subMessage').innerHTML=`<div class="notice good">Inspected ${state.subFilesMeta.length} workbook(s). Choose the sheets to read, then extract.</div>`;
  setStatus('Choose submission sheets');
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
  $('subExtractMessage').innerHTML=`<div class="notice good">Extracted <b>${j.count}</b> currency-tagged lines. Generated <b>${state.suggestions.length}</b> candidate matches.</div>`;
  $('subDownloadRow').style.display=state.subPreview.length?'flex':'none';
  table('subTable',state.subPreview,[
    {key:'submission_file',label:'File'},{key:'sheet',label:'Sheet'},{key:'source_cell',label:'Cell'},
    {key:'hierarchy_path',label:'Section'},{key:'line_description',label:'Line description'},
    {key:'currency',label:'Currency'},{key:'is_total',label:'Total row?',render:v=>v?'<span class="pill MATCH">TOTAL</span>':''},
    {key:'normalized_amount',label:'Amount'}]);
  renderSuggestions(); renderSubPicker(); renderTbPicker();
  setStatus('Submissions extracted'); go('mapping');
}
function downloadSubmissions(){
  if(!state.session){toast('No active session yet');return}
  window.location=backendUrl('/api/submissions/export?session_id='+encodeURIComponent(state.session));
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
          <div class="matchBasis muted">Section match: <b>${esc(s.match_basis.section_match||'—')}</b> · Sub-group match: <b>${esc(s.match_basis.subgroup_match||'—')}</b> · ${s.match_basis.used_total_rows?'Used pre‑computed Total rows':(isClub?`Auto-clubbed a subset of lines (excluded ${s.match_basis.excluded_lines})`:'Summed individual lines')}${sLabel?` · Scale detected: ${esc(sLabel)}`:''}</div>
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
      const accounts = ccy==='TOTAL' ? allAccounts : allAccounts.filter(a=>a.currency===ccy);
      out.push({label:label||node.name,level:node.level,currency:ccy,amount:amt,accounts});
    });
    (node.children||[]).forEach(c=>walk(c,nextPath));
  }
  (tree||[]).forEach(g=>walk(g,[]));
  return out;
}
function filteredIndexed(list,query){
  const q=(query||'').toLowerCase();
  const idx=list.map((it,i)=>({it,i}));
  return q?idx.filter(({it})=>JSON.stringify(it).toLowerCase().includes(q)):idx;
}
// merges whatever's checked at both the account level and the group
// (pivot) level into one deduped list of TB accounts, so a user can freely
// mix "this whole group" with "plus this one extra account" in one match
// without double-counting anything the group already covers.
function collectSelectedTbAccounts(){
  const map=new Map();
  [...state.tbSel].map(i=>state.tbItems[i]).filter(Boolean).forEach(it=>{
    map.set(it.account+'|'+it.currency+'|'+it.bs_mapping,it);
  });
  [...state.tbSelGroups].map(i=>state.tbGroupItems[i]).filter(Boolean).forEach(g=>{
    g.accounts.forEach(a=>map.set(a.account+'|'+a.currency+'|'+a.bs_mapping,a));
  });
  return [...map.values()];
}
function renderTbPicker(){
  const list=$('tbPickerList'); if(!list) return;
  document.querySelectorAll('#tbLevelToggle .levelBtn').forEach(b=>b.classList.toggle('active',b.dataset.level===state.tbLevel));
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
      <div class="pMeta"><div class="pTitle">${esc(it.label)} <span class="chip">${it.accounts.length} acct${it.accounts.length!==1?'s':''}</span></div><div class="pSub">BS Mapping group total (pivot level)</div></div>
      <div class="pAmt">${esc(it.currency)} ${esc(fmt(it.amount))}</div>
    </label>`:`
    <label class="pickerItem">
      <input type="checkbox" data-action="${action}" data-index="${i}" ${selSet.has(i)?'checked':''}>
      <div class="pMeta"><div class="pTitle">${esc(it.account)} — ${esc(it.account_desc)}</div><div class="pSub">${esc(it.bs_mapping)}</div></div>
      <div class="pAmt">${esc(it.currency)} ${esc(fmt(it.amount))}</div>
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
  updateMatchPreview();
}
function updateSubPickerSum(){
  const sel=[...state.subSel].map(i=>state.subPreview[i]).filter(Boolean);
  if($('subPickerCount')) $('subPickerCount').textContent=sel.length+' selected';
  if($('subPickerSum')) $('subPickerSum').textContent=fmt(sel.reduce((a,it)=>a+it.normalized_amount,0));
  updateMatchPreview();
}

// common reporting-scale multipliers to check purely from the numbers —
// mirrors bahrain_iraq_algorithm.engine.best_scale_fit. Lets the manual
// builder catch "submission is quietly divided by 1000" even when nothing
// in the sheet's header text said so.
const SCALE_CANDIDATES=[1000,1000000,0.001,0.000001];
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
  const subSel=[...state.subSel].map(i=>state.subPreview[i]).filter(Boolean);
  if(!tbSel.length||!subSel.length){ box.innerHTML=''; return; }
  const tbSum=tbSel.reduce((a,it)=>a+it.amount,0);
  const rawSubSum=subSel.reduce((a,it)=>a+it.normalized_amount,0);
  const fit=detectScale(tbSum,rawSubSum);
  if(!fit){
    const tolAbs=+($('tolAbs')?.value||1), tolPct=+($('tolPct')?.value||0.0001);
    const tol=Math.max(tolAbs,Math.abs(tbSum)*tolPct);
    const isMatch=Math.abs(tbSum-rawSubSum)<=tol;
    box.innerHTML=isMatch
      ? `<div class="notice good" style="margin:0 0 14px">✓ Values match directly — TB ${esc(fmt(tbSum))} vs submission ${esc(fmt(rawSubSum))}.</div>`
      : '';
  }else{
    box.innerHTML=`<div class="notice" style="margin:0 0 14px">⚠ TB ${esc(fmt(tbSum))} vs submission ${esc(fmt(rawSubSum))} don't match directly, but submission ${esc(scaleLabel(fit.scale))}${fit.sign<0?' (sign flipped)':''} = ${esc(fmt(rawSubSum*fit.scale*fit.sign))} — this will be applied automatically when you add the match.</div>`;
  }
}
function addManualMatch(){
  const tbSelItems=collectSelectedTbAccounts();
  if(!tbSelItems.length||!state.subSel.size){toast('Select at least one item on each side');return}
  const label=($('matchLabel').value||'').trim();
  if(!label){toast('Give this match a short description');return}
  const subSelItems=[...state.subSel].map(i=>state.subPreview[i]);
  const ccySet=new Set(tbSelItems.map(it=>it.currency));
  const currency=ccySet.size===1?[...ccySet][0]:'TOTAL';
  if(ccySet.size>1) toast('Selected TB items span multiple currencies — compared as a combined TOTAL.');
  const tbSum=tbSelItems.reduce((a,it)=>a+it.amount,0);
  const rawSubSum=subSelItems.reduce((a,it)=>a+it.normalized_amount,0);
  const fit=detectScale(tbSum,rawSubSum);
  if(fit) toast(`Detected a ${scaleLabel(fit.scale)} scale mismatch${fit.sign<0?' (and a sign flip)':''} — applied automatically.`);
  const scale=fit?fit.scale:1, signAdj=fit?fit.sign:1;
  const tb_components=tbSelItems.map(it=>({account:it.account,account_desc:it.account_desc,bs_mapping:it.bs_mapping,currency:it.currency,amount:it.amount,sign:1}));
  const components=subSelItems.map(it=>({submission_file:it.submission_file,sheet:it.sheet,row_number:it.row_number,source_cell:it.source_cell,
    line_description:it.line_description,currency:it.currency,is_total:!!it.is_total,amount:it.normalized_amount,multiplier:scale,sign:signAdj}));
  state.manualMatches.push({bs_mapping:label,label,currency,rule_type:'MANUAL_CLUB',source:'MANUAL',tb_components,components,warnings:[]});
  state.tbSel.clear(); state.tbSelGroups.clear(); state.subSel.clear(); $('matchLabel').value='';
  renderTbPicker(); renderSubPicker(); renderManualMatches(); updateMatchPreview();
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
    return `<div class="manualMatchCard ${(m.warnings&&m.warnings.length)?'hasWarning':''}">
      <div class="manualMatchHead">
        <b>${esc(m.label||m.bs_mapping)}</b>
        <span class="chip">${esc(m.currency)}</span>
        <span class="chip">${esc(m.source||'MANUAL')}</span>
        ${sLabel?`<span class="chip chipClub">Scale ${esc(sLabel)} applied</span>`:''}
        <span class="pill ${status}">${status}</span>
        <button type="button" class="danger" data-action="remove-manual-match" data-index="${i}">Remove</button>
      </div>
      <div class="matchSides">
        <div class="matchSide"><small>Trial Balance (${(m.tb_components||[]).length})</small>
          <ul>${(m.tb_components||[]).map(c=>`<li><span>${esc(c.account)} ${esc(c.account_desc)}</span><b>${esc(fmt(c.amount))}</b></li>`).join('')||'<li class="muted">none resolved</li>'}</ul>
          <div class="muted" style="margin-top:6px">Total: <b>${esc(fmt(tbSum))}</b></div>
        </div>
        <div class="matchArrow">=</div>
        <div class="matchSide"><small>Submission (${(m.components||[]).length})</small>
          <ul>${(m.components||[]).map(c=>`<li><span>${esc(c.sheet)} · ${esc(c.line_description)}</span><b>${esc(fmt(c.amount))}</b></li>`).join('')||'<li class="muted">none resolved</li>'}</ul>
          <div class="muted" style="margin-top:6px">Total: <b>${esc(fmt(subSum))}</b></div>
        </div>
      </div>
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

async function runRecon(){
  const approved=state.suggestions.filter((_,i)=>state.suggestionDecisions[i]).map(s=>({bs_mapping:s.bs_mapping,currency:s.currency,rule_type:s.rule_type,source:'AUTO',components:s.components}));
  const manual=state.manualMatches.map(m=>({bs_mapping:m.label||m.bs_mapping,currency:m.currency,rule_type:m.rule_type,source:m.source||'MANUAL',tb_components:m.tb_components,components:m.components}));
  const rules=[...approved,...manual];
  if(!rules.length){toast('Approve at least one suggestion or build a manual match');return}
  const j=await api('/api/reconcile',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({session_id:state.session,rules,tolerance_abs:+$('tolAbs').value,tolerance_pct:+$('tolPct').value})});
  state.results=j.results||[]; state.lineage=j.lineage||[];
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

function downloadOutput(){
  if(!state.session){toast('No active session yet');return}
  window.location=backendUrl('/api/download?session_id='+encodeURIComponent(state.session));
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
  'download-submissions':()=>downloadSubmissions(),
  'set-tb-level':(el)=>{ state.tbLevel=el.dataset.level; renderTbPicker(); },
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
  'remove-manual-match':(el)=>{state.manualMatches.splice(+el.dataset.index,1);renderManualMatches()},
  'download-mapping':()=>downloadMapping(),
  'apply-mapping':(el)=>guarded(el,applyMapping),
  'run-recon':(el)=>guarded(el,runRecon),
  'download':()=>downloadOutput(),
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
  if(el.dataset.action==='toggle-sheet'){
    state.selection[el.dataset.file][el.dataset.sheet].checked=el.checked; return;
  }
  if(el.dataset.action==='set-header-row'){
    state.selection[el.dataset.file][el.dataset.sheet].header_row=el.value?parseInt(el.value,10):null; return;
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
});

init();
