let state={session:null,pivot:[],pivotCols:[],tbTree:[],subFilesMeta:[],selection:{},subPreview:[],
  suggestions:[],suggestionDecisions:{},rules:[],results:[],lineage:[]};
const $=id=>document.getElementById(id);
function url(p){return typeof getWebAppBackendUrl==='function'?getWebAppBackendUrl(p):p}
async function init(){let r=await fetch(url('/api/session'),{method:'POST'});state.session=(await r.json()).session_id;setStatus('Session ready')}
function setStatus(s){$('runStatus').textContent=s}
function toast(s){let t=$('toast');t.textContent=s;t.style.display='block';setTimeout(()=>t.style.display='none',3200)}
function go(id){document.querySelectorAll('.page,.step').forEach(x=>x.classList.remove('active'));$(id).classList.add('active');document.querySelector(`.step[data-page="${id}"]`).classList.add('active')}
document.querySelectorAll('.step').forEach(x=>x.onclick=()=>go(x.dataset.page));
function fmt(v){return typeof v==='number'?v.toLocaleString(undefined,{maximumFractionDigits:2}):(v??'')}
function esc(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function table(id,rows,cols){let t=$(id);if(!rows.length){t.innerHTML='<tr><td>No records yet</td></tr>';return}
  t.innerHTML='<thead><tr>'+cols.map(c=>`<th>${c.label||c.key}</th>`).join('')+'</tr></thead><tbody>'+
  rows.map(r=>'<tr>'+cols.map(c=>`<td>${c.render?c.render(r[c.key],r):esc(fmt(r[c.key]))}</td>`).join('')+'</tr>').join('')+'</tbody>'}

/* ---------------- Trial Balance ---------------- */

async function uploadTB(){
  let f=$('tbFile').files[0]; if(!f)return toast('Choose a Trial Balance workbook');
  let fd=new FormData(); fd.append('session_id',state.session); fd.append('file',f);
  setStatus('Processing Trial Balance');
  let r=await fetch(url('/api/tb'),{method:'POST',body:fd}), j=await r.json();
  if(!j.ok){$('tbMessage').innerHTML=`<div class="notice">${esc(j.error)}</div>`;setStatus('Trial Balance failed');return}
  state.pivot=j.preview; state.pivotCols=j.columns; state.tbTree=j.tree;
  $('tbMessage').innerHTML=`<div class="notice good">Detected sheet <b>${esc(j.meta.sheet)}</b>, header row <b>${j.meta.header_row}</b>, ${j.meta.rows} data rows.</div>`;
  $('kpis').innerHTML=Object.entries(j.kpis).map(([k,v])=>`<div class="kpi"><small>${k.replace('_',' ')}</small><b>${fmt(v)}</b></div>`).join('');
  renderPivot(); renderTbTree(); setStatus('TB pivot ready'); go('pivot')
}
function renderPivot(){let q=($('pivotSearch')?.value||'').toLowerCase(),rows=state.pivot.filter(r=>JSON.stringify(r).toLowerCase().includes(q));
  table('pivotTable',rows,state.pivotCols.map(c=>({key:c,label:c})))}

function currencyChips(totals){
  if(!totals) return '';
  return Object.entries(totals).filter(([k])=>k!=='TOTAL').map(([k,v])=>`<span class="chip">${esc(k)} ${fmt(v)}</span>`).join('')+
    `<span class="chip chipTotal">TOTAL ${fmt(totals.TOTAL)}</span>`;
}
function renderTbTree(){
  let html=(state.tbTree||[]).map(g=>`
    <details class="treeNode groupNode" open>
      <summary><span class="nodeName">${esc(g.name)}</span>${currencyChips(g.currency_totals)}</summary>
      <div class="treeChildren">${(g.children||[]).map(c=>`
        <details class="treeNode subNode">
          <summary><span class="nodeName">${esc(c.name)}</span>${currencyChips(c.currency_totals)}<span class="muted">${(c.accounts||[]).length} account rows</span></summary>
          <div class="accountList">${(c.accounts||[]).map(a=>`<div class="acctRow"><span>${esc(a.account)} — ${esc(a.account_desc)}</span><span class="muted">${esc(a.currency)}</span><b>${fmt(a.amount)}</b></div>`).join('')||'<div class="muted" style="padding:8px">No account rows</div>'}</div>
        </details>`).join('')}
        ${(g.accounts||[]).length?`<div class="accountList">${g.accounts.map(a=>`<div class="acctRow"><span>${esc(a.account)} — ${esc(a.account_desc)}</span><span class="muted">${esc(a.currency)}</span><b>${fmt(a.amount)}</b></div>`).join('')}</div>`:''}
      </div>
    </details>`).join('');
  $('tbTreePreview').innerHTML=html||'<div class="muted">No structure detected yet.</div>';
}

/* ---------------- Submissions ---------------- */

async function uploadSubs(){
  let fs=[...$('subFiles').files]; if(!fs.length)return toast('Choose submission files');
  let fd=new FormData(); fd.append('session_id',state.session); fs.slice(0,3).forEach(f=>fd.append('files',f));
  setStatus('Inspecting submission workbooks');
  let r=await fetch(url('/api/submissions/upload'),{method:'POST',body:fd}), j=await r.json();
  if(!j.ok)return toast(j.error);
  state.subFilesMeta=j.files; state.selection={};
  j.files.forEach(f=>{ state.selection[f.file]={};
    f.sheets.forEach(s=>{ state.selection[f.file][s.sheet]={checked: s.header_row!=null && s.score>=3, header_row: s.header_row}; }); });
  renderSheetPicker();
  $('extractActions').style.display='flex';
  $('subMessage').innerHTML=`<div class="notice good">Inspected ${j.files.length} workbook(s). Choose the sheets to read, then extract.</div>`;
  setStatus('Choose submission sheets')
}
function toggleSheet(file,sheet,checked){state.selection[file][sheet].checked=checked}
function setHeaderRow(file,sheet,val){state.selection[file][sheet].header_row=val?parseInt(val,10):null}
function renderSheetPicker(){
  $('sheetPicker').innerHTML=state.subFilesMeta.map(f=>`
    <div class="fileCard">
      <div class="fileCardHead"><b>${esc(f.file)}</b><span class="muted">${f.sheets.length} sheet(s)</span></div>
      ${f.sheets.map(s=>{
        const sel=state.selection[f.file][s.sheet];
        const ccy=Object.entries(s.currency_columns||{}).map(([col,info])=>`<span class="chip">${col}: ${esc(info.currency)}${info.scale>1?` ×${info.scale}`:''}</span>`).join('')||'<span class="muted">No currency columns detected — will use raw numeric cells</span>';
        return `<label class="sheetRow">
          <input type="checkbox" ${sel.checked?'checked':''} onchange="toggleSheet('${esc(f.file)}','${esc(s.sheet)}',this.checked)">
          <div class="sheetInfo">
            <div class="sheetTitle">${esc(s.sheet)} <span class="muted">${s.rows}×${s.columns}</span></div>
            <div class="sheetChips">${ccy}</div>
          </div>
          <div class="sheetHeaderRow"><small>Header row</small><input type="number" min="1" value="${sel.header_row??''}" onchange="setHeaderRow('${esc(f.file)}','${esc(s.sheet)}',this.value)"></div>
        </label>`}).join('')}
    </div>`).join('');
}
async function extractSubs(){
  let selection={};
  for(const file in state.selection){
    const sheets=[];
    for(const sheet in state.selection[file]){
      const s=state.selection[file][sheet];
      if(s.checked) sheets.push(s.header_row?{sheet,header_row:s.header_row}:{sheet});
    }
    if(sheets.length) selection[file]=sheets;
  }
  if(!Object.keys(selection).length) return toast('Select at least one sheet');
  setStatus('Extracting submission sheets');
  let r=await fetch(url('/api/submissions/extract'),{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({session_id:state.session,selection})}), j=await r.json();
  if(!j.ok){$('subExtractMessage').innerHTML=`<div class="notice">${esc(j.error)}</div>`;return}
  state.subPreview=j.preview; state.suggestions=j.suggestions||[];
  state.suggestionDecisions={}; state.suggestions.forEach((s,i)=>state.suggestionDecisions[i]=s.confidence>=0.6);
  $('subExtractMessage').innerHTML=`<div class="notice good">Extracted <b>${j.count}</b> currency-tagged lines. Generated <b>${state.suggestions.length}</b> candidate matches.</div>`;
  table('subTable',state.subPreview,[
    {key:'submission_file',label:'File'},{key:'sheet',label:'Sheet'},{key:'source_cell',label:'Cell'},
    {key:'hierarchy_path',label:'Section'},{key:'line_description',label:'Line description'},
    {key:'currency',label:'Currency'},{key:'is_total',label:'Total row?',render:v=>v?'<span class="pill MATCH">TOTAL</span>':''},
    {key:'normalized_amount',label:'Amount'}]);
  renderSuggestions();
  setStatus('Submissions extracted'); go('mapping')
}

/* ---------------- Mapping Studio ---------------- */

function toggleSuggestion(i,checked){state.suggestionDecisions[i]=checked; renderSuggestionSummary()}
function renderSuggestionSummary(){
  const accepted=state.suggestions.filter((_,i)=>state.suggestionDecisions[i]);
  const groups=new Set(accepted.map(s=>s.group));
  $('suggestionSummary').innerHTML=[
    ['Suggestions',state.suggestions.length],['Approved',accepted.length],
    ['TB groups covered',groups.size],['Avg confidence',accepted.length?Math.round(100*accepted.reduce((a,s)=>a+s.confidence,0)/accepted.length)+'%':'—']
  ].map(([k,v])=>`<div class="kpi"><small>${k}</small><b>${v}</b></div>`).join('');
}
function renderSuggestions(){
  $('suggestions').innerHTML=state.suggestions.map((s,i)=>{
    const status = Math.abs(s.difference)<1e-6?'MATCH':(Math.abs(s.difference)<=1 || Math.abs(s.difference)/Math.max(Math.abs(s.tb_amount),1)<=0.0001?'MATCH_WITHIN_TOLERANCE':'REVIEW_REQUIRED');
    return `<div class="suggestionCard ${state.suggestionDecisions[i]?'accepted':''}">
      <div class="suggestionHead">
        <label class="acceptToggle"><input type="checkbox" ${state.suggestionDecisions[i]?'checked':''} onchange="toggleSuggestion(${i},this.checked)"><span>Approve</span></label>
        <div class="suggestionTitle"><b>${esc(s.bs_mapping)}</b><span class="chip">${esc(s.currency)}</span><span class="pill ${status}">${status}</span></div>
        <span class="confidence">confidence ${Math.round(s.confidence*100)}%</span>
      </div>
      <div class="suggestionBody">
        <div class="amountCols">
          <div><small>TB amount</small><b>${fmt(s.tb_amount)}</b></div>
          <div><small>Submission amount</small><b>${fmt(s.suggested_submission_amount)}</b></div>
          <div><small>Difference</small><b class="${status==='REVIEW_REQUIRED'?'bad':'good'}">${fmt(s.difference)}</b></div>
          <div><small>Sign applied</small><b>${s.sign_applied<0?'Flipped (×-1)':'As reported'}</b></div>
        </div>
        <details class="evidence"><summary>Match basis &amp; evidence (${s.components.length} line${s.components.length!==1?'s':''})</summary>
          <div class="matchBasis muted">Section match: <b>${esc(s.match_basis.section_match||'—')}</b> · Sub-group match: <b>${esc(s.match_basis.subgroup_match||'—')}</b> · ${s.match_basis.used_total_rows?'Used pre‑computed Total rows':'Summed individual lines'}</div>
          <table class="miniTable"><thead><tr><th>File</th><th>Sheet</th><th>Cell</th><th>Description</th><th>Amount</th></tr></thead>
          <tbody>${s.components.map(c=>`<tr><td>${esc(c.submission_file)}</td><td>${esc(c.sheet)}</td><td>${esc(c.source_cell||c.row_number)}</td><td>${esc(c.line_description)}</td><td>${fmt(c.amount)}</td></tr>`).join('')}</tbody></table>
        </details>
      </div>
    </div>`}).join('') || '<div class="notice">No candidate matches were generated. Add manual rules below.</div>';
  renderSuggestionSummary();
}

function addRule(){state.rules.push({bs_mapping:'',currency:'TOTAL',rule_type:'DIRECT',components:[{submission_file:'',sheet:'',row_number:'',currency:'',multiplier:1,sign:1}]});renderRules()}
function renderRules(){
  $('rules').innerHTML=state.rules.map((r,i)=>`<div class="rule">
    <label>BS Mapping (as in TB)<input value="${esc(r.bs_mapping)}" onchange="state.rules[${i}].bs_mapping=this.value"></label>
    <label>Currency<input value="${esc(r.currency)}" placeholder="TOTAL / USD / IQD" onchange="state.rules[${i}].currency=this.value"></label>
    <label>Rule type<select onchange="state.rules[${i}].rule_type=this.value"><option>DIRECT</option><option>ONE_TO_MANY</option><option>MANY_TO_ONE</option><option>COMPLEX</option></select></label>
    <label>Submission components<div>${r.components.map((c,k)=>`<div class="component">
      <input placeholder="File name" value="${esc(c.submission_file)}" onchange="state.rules[${i}].components[${k}].submission_file=this.value">
      <input placeholder="Sheet" value="${esc(c.sheet)}" onchange="state.rules[${i}].components[${k}].sheet=this.value">
      <input placeholder="Row" type="number" value="${c.row_number}" onchange="state.rules[${i}].components[${k}].row_number=this.value">
      <input placeholder="Ccy" value="${esc(c.currency)}" onchange="state.rules[${i}].components[${k}].currency=this.value">
      <select onchange="state.rules[${i}].components[${k}].sign=+this.value"><option value="1">Add</option><option value="-1">Subtract</option></select>
      <button class="danger" onclick="state.rules[${i}].components.splice(${k},1);renderRules()">×</button></div>`).join('')}
      <button class="secondary" onclick="state.rules[${i}].components.push({submission_file:'',sheet:'',row_number:'',currency:'',multiplier:1,sign:1});renderRules()">+ component</button></div></label>
    <button class="danger" onclick="state.rules.splice(${i},1);renderRules()">Remove</button>
  </div>`).join('');
}

async function runRecon(){
  const approved=state.suggestions.filter((_,i)=>state.suggestionDecisions[i]).map(s=>({bs_mapping:s.bs_mapping,currency:s.currency,rule_type:s.rule_type,source:'AUTO',components:s.components}));
  const manual=state.rules.filter(r=>r.bs_mapping).map(r=>({...r,source:'MANUAL'}));
  const rules=[...approved,...manual];
  if(!rules.length) return toast('Approve at least one suggestion or add a manual rule');
  let r=await fetch(url('/api/reconcile'),{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({session_id:state.session,rules,tolerance_abs:+$('tolAbs').value,tolerance_pct:+$('tolPct').value})}), j=await r.json();
  if(!j.ok)return toast(j.error);
  state.results=j.results; state.lineage=j.lineage;
  let counts={}; j.results.forEach(x=>counts[x.status]=(counts[x.status]||0)+1);
  $('resultSummary').innerHTML=Object.entries(counts).map(([k,v])=>`<div class="kpi"><small>${k}</small><b>${v}</b></div>`).join('');
  table('resultTable',j.results,[
    {key:'bs_mapping',label:'BS Mapping'},{key:'currency',label:'Currency'},
    {key:'tb_amount',label:'TB Amount'},{key:'submission_amount',label:'Submission Amount'},{key:'difference',label:'Difference'},
    {key:'variance_pct',label:'Variance %',render:v=>(v*100).toFixed(4)+'%'},
    {key:'status',label:'Status',render:v=>`<span class="pill ${v}">${v}</span>`},
    {key:'rule_type',label:'Rule'},{key:'source',label:'Source'}]);
  renderLineageTree();
  setStatus('Reconciliation complete'); go('results')
}

/* ---------------- Reconciliation Map ---------------- */

function statusPill(status){return status?`<span class="pill ${status}">${status}</span>`:''}
function renderMatchRow(m){
  let evidence=[]; try{evidence=JSON.parse(m.evidence||'[]')}catch(e){}
  return `<details class="matchRow">
    <summary><span class="chip">${esc(m.currency)}</span>${statusPill(m.status)}
      <span class="muted">TB ${fmt(m.tb_amount)} vs submission ${fmt(m.submission_amount)} (Δ ${fmt(m.difference)})</span>
      <span class="chip">${esc(m.source||'')}</span></summary>
    <table class="miniTable"><thead><tr><th>File</th><th>Sheet</th><th>Row</th><th>Description</th><th>Amount</th><th>Currency</th></tr></thead>
    <tbody>${evidence.map(e=>`<tr><td>${esc(e.file)}</td><td>${esc(e.sheet)}</td><td>${e.row}</td><td>${esc(e.description)}</td><td>${fmt(e.amount)}</td><td>${esc(e.currency||'')}</td></tr>`).join('')||'<tr><td colspan="6" class="muted">No evidence rows</td></tr>'}</tbody></table>
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
      ${accounts.length?`<details class="accountsDetail"><summary class="muted">${accounts.length} TB account row(s)</summary><div class="accountList">${accounts.map(a=>`<div class="acctRow"><span>${esc(a.account)} — ${esc(a.account_desc)}</span><span class="muted">${esc(a.currency)}</span><b>${fmt(a.amount)}</b></div>`).join('')}</div></details>`:''}
      ${children.map(renderLineageNode).join('')}
    </div>
  </details>`;
}
function renderLineageTree(){
  $('mapLegend').innerHTML=['MATCH','MATCH_WITHIN_TOLERANCE','REVIEW_REQUIRED'].map(s=>`<span class="pill ${s}">${s}</span>`).join(' ');
  $('lineageTree').innerHTML=(state.lineage||[]).map(renderLineageNode).join('') || '<div class="muted">Run the reconciliation to build the map.</div>';
}

function downloadOutput(){window.location=url('/api/download?session_id='+encodeURIComponent(state.session))}
init();
