import os, sys, json, uuid, tempfile, traceback
from pathlib import Path
from flask import request, jsonify, Response

HERE=os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
if HERE not in sys.path: sys.path.insert(0,HERE)
from bahrain_iraq_algorithm import ReconEngine
# depth search lives in its own module; if the project library is older than this backend, keep the
# rest of the app working and report a clear message from the depth-search endpoints instead
try:
    from bahrain_iraq_algorithm import depth_search as DS
    DS_ERROR=None
except Exception as _e:
    DS=None; DS_ERROR=str(_e)

try:
    from bahrain_iraq_algorithm import translation as TR
    TR_ERROR=None
except Exception as _e:
    TR=None; TR_ERROR=str(_e)
try:
    from bahrain_iraq_algorithm import outstanding_report as OR
    OR_ERROR=None
except Exception as _e:
    OR=None; OR_ERROR=str(_e)

try:
    from bahrain_iraq_algorithm import default_report as DR
    DR_ERROR=None
except Exception as _e:
    DR=None; DR_ERROR=str(_e)

SESSIONS={}

def get_session(sid=None):
    sid=sid or str(uuid.uuid4())
    d=SESSIONS.setdefault(sid,{'dir':tempfile.mkdtemp(prefix='bahrain_iraq_'),'tb_path':None,'tb':None,'pivot':None,
                               'tb_tree':None,'sub_files':[],'sub_inspect':{},'sub_df':None,'suggestions':None,
                               'rules':[],'recon':None,'lineage':None})
    return sid,d

# send_file's filename kwarg was renamed (attachment_filename -> download_name) in Flask 2.0,
# so build the attachment response by hand to work on any Flask version Dataiku ships.
def xlsx_response(path,filename):
    with open(path,'rb') as fh: data=fh.read()
    return bytes_response(data,filename,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

def bytes_response(data,filename,mimetype):
    return Response(data,mimetype=mimetype,
                    headers={'Content-Disposition':'attachment; filename="%s"'%filename,'Content-Length':str(len(data))})

def save_upload(f,d,prefix):
    name=Path(f.filename or prefix+'.xlsx').name; path=os.path.join(d['dir'],prefix+'_'+name); f.save(path); return path

# --- make sure every response is JSON, never Flask's default HTML error page,
#     and never crashes the whole request silently on the frontend ---

@app.errorhandler(Exception)
def _handle_any_error(e):
    traceback.print_exc()
    return jsonify({'ok':False,'error':str(e) or e.__class__.__name__}),500

@app.errorhandler(404)
def _handle_404(e):
    return jsonify({'ok':False,'error':f'No such API route: {request.path}'}),404

@app.after_request
def _cors(resp):
    # defensive: harmless if the webapp is served same-origin (the normal
    # Dataiku case), but prevents a silent CORS failure if it isn't.
    resp.headers.setdefault('Access-Control-Allow-Origin','*')
    resp.headers.setdefault('Access-Control-Allow-Headers','Content-Type')
    resp.headers.setdefault('Access-Control-Allow-Methods','GET,POST,OPTIONS')
    return resp

@app.route('/api/session',methods=['POST','OPTIONS'])
def session():
    sid,_=get_session(); return jsonify({'ok':True,'session_id':sid})

@app.route('/api/inspect',methods=['POST','OPTIONS'])
def inspect():
    sid,d=get_session(request.form.get('session_id')); f=request.files.get('file')
    if not f:return jsonify({'ok':False,'error':'No file uploaded'}),400
    path=save_upload(f,d,'inspect'); return jsonify({'ok':True,'session_id':sid,'file':Path(f.filename).name,'sheets':ReconEngine().inspect_workbook(path)})

@app.route('/api/tb',methods=['POST','OPTIONS'])
def upload_tb():
    sid,d=get_session(request.form.get('session_id')); f=request.files.get('file')
    if not f:return jsonify({'ok':False,'error':'Upload a Trial Balance workbook.'}),400
    path=save_upload(f,d,'tb'); eng=ReconEngine(); tb,meta=eng.load_tb(path,request.form.get('sheet') or None,request.form.get('header_row') or None)
    pivot=eng.make_pivot(tb); tree=eng.build_tb_tree(tb)
    d.update(tb_path=path,tb=tb,pivot=pivot,tb_tree=tree)
    return jsonify({'ok':True,'session_id':sid,'meta':meta,'local_currency':getattr(eng,'LOCAL_CCY','IQD'),'engine_has_frx':hasattr(eng,'FRX_CODE'),
                    'kpis':{'rows':len(tb),'groups':int(tb.bs_mapping.nunique()),'currencies':int(tb.tran_ccy.nunique()),
                            'balance':float(tb.adjusted_balance.sum()),'unmapped':int((tb.bs_mapping=='').sum())},
                    'preview':json.loads(pivot.head(20000).to_json(orient='records')),'columns':pivot.columns.tolist(),
                    'tree':tree})

@app.route('/api/submissions/upload',methods=['POST','OPTIONS'])
def submissions_upload():
    sid,d=get_session(request.form.get('session_id')); fs=request.files.getlist('files')
    if not fs:return jsonify({'ok':False,'error':'Upload at least one submission workbook.'}),400
    eng=ReconEngine(); d['sub_files']=[]; d['sub_inspect']={}
    out=[]
    d['translation']=None; d['work_language']=None
    for i,f in enumerate(fs[:8],1):
        label=Path(f.filename or '').name
        if Path(label).suffix.lower() not in ('.xlsx','.xlsm','.xls','.xlsb'):
            return jsonify({'ok':False,'error':'%s is not an Excel workbook (.xlsx, .xlsm, .xls or .xlsb).'%label}),400
        path=save_upload(f,d,f'sub{i}')
        if Path(label).suffix.lower() in ('.xls','.xlsb'):
            if TR is None: return jsonify({'ok':False,'error':'Reading .xls/.xlsb needs the latest bahrain_iraq_algorithm library. Detail: %s'%TR_ERROR}),400
            path=TR.to_xlsx(path,d['dir'])          # the rest of the app reads .xlsx
        d['sub_files'].append({'label':label,'path':path})
        sheets=eng.inspect_workbook(path)
        d['sub_inspect'][label]=sheets
        guessed=DS.guess_file_type(label,[s['sheet'] for s in sheets]) if DS else 'any'
        ar=TR.count_arabic(path) if TR else {'cells':0,'sheets':[],'sheet_names':[],'sample':[]}
        out.append({'file':label,'sheets':sheets,'guessed_type':guessed,'arabic_cells':ar['cells'],
                    'arabic_sheets':ar['sheets'],'arabic_sample':ar['sample']})
    d['depth']=None
    has_ar=bool(TR) and any(o['arabic_cells'] or any(TR.contains_arabic(x['sheet']) for x in o['sheets']) for o in out)
    return jsonify({'ok':True,'files':out,'has_arabic':has_ar})

@app.route('/api/submissions/extract',methods=['POST','OPTIONS'])
def submissions_extract():
    body=request.get_json(force=True,silent=True) or {}
    sid,d=get_session(body.get('session_id'))
    if not d['sub_files']:return jsonify({'ok':False,'error':'Upload submission workbooks first.'}),400
    selection=body.get('selection',{})
    eng=ReconEngine(); frames=[]
    for entry in d['sub_files']:
        label,path=entry['label'],entry['path']
        sheets=selection.get(label)
        if not sheets: continue
        frames.append(eng.extract_submission(path,label,sheets=sheets))
    import pandas as pd
    d['sub_df']=pd.concat(frames,ignore_index=True) if frames else pd.DataFrame()
    d['suggestions']=eng.auto_match(d['tb_tree'] or [], d['sub_df']) if d['tb_tree'] is not None else []

    # a bundled, name/description-driven mapping (never cell addresses) is
    # resolved automatically here too, so a known chart of accounts is part
    # of reconciliation by default — no manual upload needed. Every attempted
    # rule is recorded (fulfilled or not, and why) so coverage is fully
    # visible rather than collapsed into a single skipped-count.
    default_matches=[]; coverage=[]; default_payload=None
    d['default_run']=None
    if d['tb'] is not None and len(d['sub_df']):
        default_matches,coverage,default_payload=_run_default(d,eng)
    d['mapping_coverage']=coverage

    preview=d['sub_df'].copy()
    if 'hierarchy' in preview.columns: preview=preview.drop(columns=['hierarchy'])
    summary={'total':len(coverage),'fulfilled':sum(1 for e in coverage if e['fulfilled']),
             'unresolved':sum(1 for e in coverage if not e['fulfilled']),
             'tb_only':sum(1 for e in coverage if e['tb_found'] and not e['sub_found']),
             'sub_only':sum(1 for e in coverage if e['sub_found'] and not e['tb_found']),
             'neither':sum(1 for e in coverage if not e['tb_found'] and not e['sub_found'])}
    # the manual match builder and amount search both need the FULL set of
    # extracted lines to be useful — a 400-row cap here silently hid every
    # line past that point (often an entire second/third uploaded file) from
    # the picker even though it was correctly extracted underneath. 20,000
    # is a safety net against a truly pathological file, not a real limit.
    return jsonify({'ok':True,'count':len(d['sub_df']),'preview':json.loads(preview.head(20000).to_json(orient='records')),
                    'suggestions':d['suggestions'],'default_mapping_matches':default_matches,
                    'mapping_coverage':coverage,'mapping_coverage_summary':summary,'default_results':default_payload})

@app.route('/api/submissions/search-amount',methods=['POST','OPTIONS'])
def submissions_search_amount():
    body=request.get_json(force=True,silent=True) or {}
    sid,d=get_session(body.get('session_id'))
    if d.get('sub_df') is None or not len(d['sub_df']):
        return jsonify({'ok':False,'error':'Process submissions first.'}),400
    amount=body.get('amount')
    if amount is None:return jsonify({'ok':False,'error':'No amount given.'}),400
    eng=ReconEngine()
    results=eng.search_amount(d['sub_df'],amount,body.get('tolerance_abs',1),body.get('tolerance_pct',0.0001),currency=body.get('currency') or None)
    return jsonify({'ok':True,'results':results[:200],'total_found':len(results)})

@app.route('/api/submissions/export',methods=['GET'])
def submissions_export():
    sid=request.args.get('session_id'); _,d=get_session(sid)
    if d.get('sub_df') is None or not len(d['sub_df']):return jsonify({'ok':False,'error':'No submission lines extracted yet.'}),400
    eng=ReconEngine()
    path=os.path.join(d['dir'],'Submissions_Simplified.xlsx')
    eng.export_submissions(d['sub_df'],path)
    return xlsx_response(path,'Submissions_Simplified.xlsx')

def _strip_placeholders(resolved):
    r=dict(resolved)
    r['tb_components']=[c for c in (resolved.get('tb_components') or []) if not c.get('placeholder')]
    r['components']=[c for c in (resolved.get('components') or []) if not c.get('placeholder')]
    return r

def _run_default(d,eng):
    """Run the bundled default mapping. Returns (matches for the Mapping Studio, coverage rows, payload for the page)."""
    matches=[]; coverage=[]
    if DR is None:
        # an older library without default_report: fall back to the plain loop so the rest of the app keeps working
        for tpl in eng.load_default_mapping():
            resolved,warnings=eng.resolve_rule_template(tpl,d['tb'],d['sub_df'])
            tb_n=len(resolved.get('tb_components') or []); sub_n=len(resolved.get('components') or [])
            ok=tb_n>0 and sub_n>0
            e={'label':resolved.get('label'),'currency':resolved.get('currency'),'rule_type':resolved.get('rule_type'),
               'fulfilled':ok,'tb_found':tb_n>0,'sub_found':sub_n>0,'tb_accounts':tb_n,'sub_lines':sub_n,'warnings':warnings,'reconciled_status':None}
            if ok:
                resolved['source']='DEFAULT'; matches.append({'resolved':resolved,'warnings':warnings})
                e['tb_amount']=sum(c['amount']*c.get('sign',1) for c in resolved['tb_components'])
                e['submission_amount']=sum(c['amount']*c.get('sign',1)*c.get('multiplier',1) for c in resolved['components'])
                e['difference']=e['tb_amount']-e['submission_amount']
            coverage.append(e)
        return matches,coverage,{'ok':False,'error':'The default-mapping report needs the latest bahrain_iraq_algorithm library '
                                 '(copy default_report.py, pdf_simple.py and engine.py into it, then restart the backend). Detail: %s'%DR_ERROR}
    run=DR.run_default_mapping(eng,d['tb'],d['sub_df'],outstanding=(d.get('outstanding') or {}).get('result'),
                               ob_tb=bool((d.get('default_opts') or {}).get('ob_tb')),col_search=(d.get('default_opts') or {}).get('col_search',True))
    d['default_run']=run
    for r in run['results']:
        if r.get('ob'): continue                 # off-balance checks live on the Default Mapping page, not in Mapping Studio
        resolved=_strip_placeholders(r['resolved'])
        ok=r['n_tb']>0 and r['n_sub']>0
        e={'label':r['label'],'currency':r['currency'],'rule_type':r['rule_type'],'fulfilled':ok,'tb_found':r['n_tb']>0,
           'sub_found':r['n_sub']>0,'tb_accounts':r['n_tb'],'sub_lines':r['n_sub'],'warnings':r['warnings'],'reconciled_status':None,
           'default_status':r['status'],'variance':r['variance']}
        if ok:
            resolved['source']='DEFAULT'; matches.append({'resolved':resolved,'warnings':r['warnings']})
            e['tb_amount']=r['tb_amount']; e['submission_amount']=r['sub_amount']; e['difference']=r['variance']
        coverage.append(e)
    return matches,coverage,_default_payload(d,run)

def _default_payload(d,run):
    files=[]
    for entry in d['sub_files']:
        sheets=[s['sheet'] for s in d['sub_inspect'].get(entry['label'],[])]
        mine=DR._results_for_file(run['results'],entry['label'],sheets)
        sub=d.get('sub_df')
        read=sorted(set(sub[sub.submission_file==entry['label']].sheet)) if (sub is not None and len(sub)) else []
        files.append({'file':entry['label'],'summary':DR.summarise(mine),'rules':[r['n'] for r in mine],'sheets_read':read})
    return {'ok':True,'results':DR.public_results(run),'summary':run['summary'],'params':run['params'],'files':files,'groups':run.get('groups',[]),'notes':run.get('notes',[]),
            'options':{'ob_tb':bool((d.get('default_opts') or {}).get('ob_tb')),'col_search':(d.get('default_opts') or {}).get('col_search',True)},
            'ob_status':((d.get('outstanding') or {}).get('status'))}

def _dm_unavailable():
    return jsonify({'ok':False,'error':'The default-mapping report needs the latest bahrain_iraq_algorithm library in your project '
                    '(copy default_report.py, pdf_simple.py and engine.py into it, then restart the backend). Detail: %s'%DR_ERROR}),400

def _dm_ready(d):
    if d.get('tb') is None: return jsonify({'ok':False,'error':'Upload and process the Trial Balance first.'}),400
    if d.get('sub_df') is None or not len(d['sub_df']): return jsonify({'ok':False,'error':'Process the submissions first.'}),400
    return None

@app.route('/api/default/run',methods=['POST','OPTIONS'])
def default_run():
    if DR is None: return _dm_unavailable()
    body=request.get_json(force=True,silent=True) or {}
    sid,d=get_session(body.get('session_id'))
    bad=_dm_ready(d)
    if bad: return bad
    eng=ReconEngine(float(body.get('tolerance_abs',1)),float(body.get('tolerance_pct',0.0001)))
    opts=dict(d.get('default_opts') or {})
    if 'ob_tb' in body: opts['ob_tb']=bool(body.get('ob_tb'))
    if 'col_search' in body: opts['col_search']=bool(body.get('col_search'))
    d['default_opts']=opts
    run=DR.run_default_mapping(eng,d['tb'],d['sub_df'],outstanding=(d.get('outstanding') or {}).get('result'),
                               ob_tb=bool(opts.get('ob_tb')),col_search=opts.get('col_search',True))
    d['default_run']=run
    return jsonify(_default_payload(d,run))

@app.route('/api/default/results',methods=['GET'])
def default_results():
    if DR is None: return _dm_unavailable()
    _,d=get_session(request.args.get('session_id'))
    if not d.get('default_run'): return jsonify({'ok':False,'error':'Process the submissions first.'}),400
    return jsonify(_default_payload(d,d['default_run']))

def _dm_out_path(d,label):
    ext='.xlsm' if label.lower().endswith('.xlsm') else '.xlsx'
    safe=''.join(ch if ch.isalnum() or ch in '-_.' else '_' for ch in Path(label).stem)
    return os.path.join(d['dir'],'default_'+safe+ext), safe+'_DefaultMapping'+ext

@app.route('/api/default/download',methods=['GET'])
def default_download():
    if DR is None: return _dm_unavailable()
    _,d=get_session(request.args.get('session_id')); label=request.args.get('file','')
    run=d.get('default_run')
    if not run: return jsonify({'ok':False,'error':'Process the submissions first.'}),400
    src=next((x for x in d['sub_files'] if x['label']==label),None)
    if not src: return jsonify({'ok':False,'error':'That file is not part of this session.'}),400
    out,name=_dm_out_path(d,label)
    DR.write_default_workbook(src['path'],out,label,d['tb'],run)
    with open(out,'rb') as fh: data=fh.read()
    return bytes_response(data,name,_depth_mime(out))

@app.route('/api/default/download-all',methods=['GET'])
def default_download_all():
    if DR is None: return _dm_unavailable()
    _,d=get_session(request.args.get('session_id'))
    run=d.get('default_run')
    if not run: return jsonify({'ok':False,'error':'Process the submissions first.'}),400
    entries=[]
    for src in d['sub_files']:
        if not any(c.get('submission_file')==src['label'] for r in run['results'] for c in r['components']): continue
        out,name=_dm_out_path(d,src['label'])
        DR.write_default_workbook(src['path'],out,src['label'],d['tb'],run)
        entries.append((name,out))
    if not entries: return jsonify({'ok':False,'error':'No submission file has a default-mapping result.'}),400
    return bytes_response(DR.build_zip(entries),'Default_Mapping_Outputs.zip','application/zip')

@app.route('/api/default/pdf',methods=['POST','OPTIONS'])
def default_pdf():
    if DR is None: return _dm_unavailable()
    body=request.get_json(force=True,silent=True) or {}
    _,d=get_session(body.get('session_id'))
    run=d.get('default_run')
    if not run: return jsonify({'ok':False,'error':'Process the submissions first.'}),400
    files=[]
    for entry in d['sub_files']:
        sheets=[x['sheet'] for x in d['sub_inspect'].get(entry['label'],[])]
        files.append({'file':entry['label'],'sheets':sheets,
                      'type_label':next((f.get('type_label','') for f in ((d.get('depth') or {}).get('files') or []) if f['file']==entry['label']),'')})
    tb_name=Path(d['tb_path']).name.split('_',1)[-1] if d.get('tb_path') else ''
    data=DR.build_pdf(run,files=files,tb_name=tb_name,suggestions=d.get('suggestions') or [],
                      approved=body.get('approved'),manual_matches=body.get('manual_matches') or [])
    return bytes_response(data,'Default_Mapping_Results.pdf','application/pdf')

@app.route('/api/rules/resolve',methods=['POST','OPTIONS'])
def rules_resolve():
    body=request.get_json(force=True,silent=True) or {}
    sid,d=get_session(body.get('session_id'))
    if d.get('tb') is None or d.get('sub_df') is None:
        return jsonify({'ok':False,'error':'Upload a Trial Balance and process submissions before importing a mapping.'}),400
    templates=body.get('templates',[])
    if not templates:return jsonify({'ok':False,'error':'The uploaded mapping file has no matches in it.'}),400
    eng=ReconEngine(); out=[]
    if DR is not None: templates=DR.expand_templates(templates,d['tb'],d['sub_df'])
    for t in templates:
        resolved,warnings=eng.resolve_rule_template(t,d['tb'],d['sub_df'])
        out.append({'resolved':resolved,'warnings':warnings})
    return jsonify({'ok':True,'results':out})

@app.route('/api/reconcile',methods=['POST','OPTIONS'])
def reconcile():
    body=request.get_json(force=True,silent=True) or {}
    sid,d=get_session(body.get('session_id'))
    if d['tb'] is None or d['sub_df'] is None:return jsonify({'ok':False,'error':'Process the Trial Balance and submissions first.'}),400
    rules=body.get('rules',[])
    eng=ReconEngine(body.get('tolerance_abs',1),body.get('tolerance_pct',0.0001))
    d['rules']=rules
    d['recon']=eng.reconcile(d['tb'],d['sub_df'],rules)
    recon_rows=json.loads(d['recon'].to_json(orient='records'))
    lineage=eng.build_lineage(d['tb_tree'] or [], recon_rows)
    d['lineage']=lineage

    # fold the actual reconciliation outcome back into the mapping coverage
    # report, so "fulfilled" (data was found) and "reconciled" (it was
    # approved AND the numbers actually matched) are visibly two different
    # questions with two different answers.
    if d.get('mapping_coverage'):
        by_key={(r['bs_mapping'],r['currency']):r for r in recon_rows}
        for entry in d['mapping_coverage']:
            r=by_key.get((entry['label'],entry['currency']))
            if r:
                entry['reconciled_status']=r['status']
                entry['tb_amount']=r['tb_amount']; entry['submission_amount']=r['submission_amount']; entry['difference']=r['difference']
            elif entry['fulfilled']:
                entry['reconciled_status']='NOT_APPROVED'

    return jsonify({'ok':True,'results':recon_rows,'lineage':lineage,'mapping_coverage':d.get('mapping_coverage',[])})

@app.route('/api/lineage',methods=['GET'])
def lineage():
    sid=request.args.get('session_id'); _,d=get_session(sid)
    if not d.get('lineage'):return jsonify({'ok':False,'error':'Run reconciliation first.'}),400
    return jsonify({'ok':True,'lineage':d['lineage']})

@app.route('/api/download',methods=['GET'])
def download():
    sid=request.args.get('session_id'); _,d=get_session(sid)
    if d['tb'] is None:return jsonify({'ok':False,'error':'No processed Trial Balance.'}),400
    eng=ReconEngine()
    flat_lineage=eng.flatten_lineage(d['lineage']) if d.get('lineage') else None
    path=os.path.join(d['dir'],'Bahrain_Iraq_Reconciliation_Output.xlsx')
    eng.export(d['tb'],d['pivot'],d['sub_df'],d['recon'],d.get('suggestions'),flat_lineage,path,mapping_coverage=d.get('mapping_coverage'))
    return xlsx_response(path,'Bahrain_Iraq_Reconciliation_Output.xlsx')


# ---------------- Depth search: TB pivot values -> every non-zero cell of every sheet ----------------

def _depth_unavailable():
    return jsonify({'ok':False,'error':'Depth search needs the latest bahrain_iraq_algorithm library in your project '
                    '(copy depth_search.py and engine.py into it, then restart the backend). Detail: %s'%DS_ERROR}),400

def _depth_out_path(d,label):
    ext='.xlsm' if label.lower().endswith('.xlsm') else '.xlsx'
    safe=''.join(ch if ch.isalnum() or ch in '-_.' else '_' for ch in Path(label).stem)
    return os.path.join(d['dir'],'depth_'+safe+ext), safe+'_DepthSearch'+ext

def _depth_mime(path):
    return 'application/vnd.ms-excel.sheet.macroEnabled.12' if path.lower().endswith('.xlsm') else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

@app.route('/api/depth-search/run',methods=['POST','OPTIONS'])
def depth_run():
    if DS is None: return _depth_unavailable()
    body=request.get_json(force=True,silent=True) or {}
    sid,d=get_session(body.get('session_id'))
    if d['tb'] is None:return jsonify({'ok':False,'error':'Upload and process the Trial Balance first.'}),400
    if not d['sub_files']:return jsonify({'ok':False,'error':'Upload the submission workbooks first (Submissions step).'}),400
    ps=body.get('params') or {}
    params={'min_value':float(ps.get('min_value',1000)),'min_scaled':float(ps.get('min_scaled',100000)),
            'tol_abs':float(ps.get('tol_abs',1)),'allow_scale':bool(ps.get('allow_scale',True)),
            'allow_sign':bool(ps.get('allow_sign',True)),'include_accounts':bool(ps.get('include_accounts',False)),
            'include_tb_ob':bool(ps.get('include_tb_ob',False))}
    if ps.get('maturity_keywords'): params['maturity_keywords']=str(ps['maturity_keywords'])
    externals=OR.depth_specs(d['outstanding']['result']) if (OR is not None and d.get('outstanding')) else None
    res=DS.run_depth_search(d['tb'],d['sub_files'],body.get('file_types') or {},params,externals=externals)
    d['depth']=res
    tmap={t['id']:t for t in res['targets']}
    files=[]
    for f in res['files']:
        files.append({'file':f['file'],'type':f['type'],'type_label':f['type_label'],'skipped':f['skipped'],'warnings':f['warnings'],
                      'stats':f['stats'],'hits':f['hits'][:3000],'hits_total':len(f['hits']),
                      'unmatched':[{'id':i,'label':tmap[i]['label'],'level':tmap[i]['level'],'currency':tmap[i]['currency'],
                                    'amount':tmap[i]['amount']} for i in f['unmatched']][:3000]})
    targets={t['id']:{'label':t['label'],'level':t['level'],'currency':t['currency'],'amount':t['amount'],'cell':t['cell'],
                      'sheet':t['sheet'],'family':t['family']} for t in res['targets']}
    ext=[{'sheet':e['sheet'],'scope':e['scope'],'category':e.get('category')} for e in res.get('externals',[])]
    return jsonify({'ok':True,'files':files,'targets':targets,'params':res['params'],'externals':ext,'notes':res.get('notes',[])})

@app.route('/api/depth-search/download',methods=['GET'])
def depth_download():
    if DS is None: return _depth_unavailable()
    sid=request.args.get('session_id'); _,d=get_session(sid); label=request.args.get('file','')
    res=d.get('depth')
    if not res:return jsonify({'ok':False,'error':'Run the depth search first.'}),400
    entry=next((f for f in res['files'] if f['file']==label),None)
    src=next((s for s in d['sub_files'] if s['label']==label),None)
    if not entry or not src:return jsonify({'ok':False,'error':'That file was not part of the last depth search.'}),400
    if entry['skipped']:return jsonify({'ok':False,'error':'That file was skipped: %s'%'; '.join(entry['warnings'])}),400
    out,name=_depth_out_path(d,label)
    DS.write_annotated_workbook(src['path'],out,label,d['tb'],res,entry)
    with open(out,'rb') as fh: data=fh.read()
    return bytes_response(data,name,_depth_mime(out))

@app.route('/api/depth-search/download-all',methods=['GET'])
def depth_download_all():
    if DS is None: return _depth_unavailable()
    sid=request.args.get('session_id'); _,d=get_session(sid)
    res=d.get('depth')
    if not res:return jsonify({'ok':False,'error':'Run the depth search first.'}),400
    entries=[]
    for entry in res['files']:
        if entry['skipped']: continue
        src=next((s for s in d['sub_files'] if s['label']==entry['file']),None)
        if not src: continue
        out,name=_depth_out_path(d,entry['file'])
        DS.write_annotated_workbook(src['path'],out,entry['file'],d['tb'],res,entry)
        entries.append((name,out))
    if not entries:return jsonify({'ok':False,'error':'No searched files to download.'}),400
    return bytes_response(DS.build_zip(entries),'Depth_Search_Outputs.zip','application/zip')


# ---------------- Arabic submissions: translate with the user's dictionary, or keep working in Arabic ----------------

def _tr_unavailable():
    return jsonify({'ok':False,'error':'Translation needs the latest bahrain_iraq_algorithm library in your project '
                    '(copy translation.py and engine.py into it, then restart the backend). Detail: %s'%TR_ERROR}),400

def _file_meta(d,eng):
    out=[]
    for sf in d['sub_files']:
        sheets=eng.inspect_workbook(sf['path']); d['sub_inspect'][sf['label']]=sheets
        ar=TR.count_arabic(sf['path'])
        out.append({'file':sf['label'],'sheets':sheets,'arabic_cells':ar['cells'],'arabic_sheets':ar['sheets'],'arabic_sample':ar['sample'],
                    'guessed_type':DS.guess_file_type(sf['label'],[x['sheet'] for x in sheets]) if DS else 'any'})
    return out

@app.route('/api/submissions/translate',methods=['POST','OPTIONS'])
def submissions_translate():
    if TR is None: return _tr_unavailable()
    sid,d=get_session(request.form.get('session_id')); f=request.files.get('dictionary')
    if not d['sub_files']:return jsonify({'ok':False,'error':'Upload the submission workbooks first.'}),400
    if not f:return jsonify({'ok':False,'error':'Upload your dictionary workbook (Arabic in column A, English in column B, header in row 1).'}),400
    dpath=save_upload(f,d,'dictionary')
    try: dictionary=TR.load_dictionary(dpath)
    except Exception as exc: return jsonify({'ok':False,'error':'Could not read the dictionary workbook: %s'%exc}),400
    if not dictionary:return jsonify({'ok':False,'error':'The dictionary has no usable rows (Arabic in column A, English in column B, header in row 1).'}),400
    eng=ReconEngine(); stats=[]; missing=set()
    for sf in d['sub_files']:
        base=sf.get('orig_path') or sf['path']; label=sf['label']
        ar=TR.count_arabic(base)
        if not ar['cells'] and not ar['sheet_names']:
            stats.append({'file':label,'translated_cells':0,'translated_sheets':0,'missing_terms':0,'skipped':True}); continue
        ext='.xlsm' if base.lower().endswith('.xlsm') else '.xlsx'
        out=os.path.join(d['dir'],'translated_'+Path(base).stem+ext)
        st=TR.translate_workbook(base,dictionary,out)
        sf['orig_path']=base; sf['path']=out; missing|=st['missing_terms']
        stats.append({'file':label,'translated_cells':st['translated_cells'],'translated_sheets':st['translated_sheets'],
                      'missing_terms':len(st['missing_terms']),'skipped':False})
    missing_path=os.path.join(d['dir'],'Missing_Arabic_Words.xlsx'); TR.write_missing_terms(missing,missing_path)
    d['translation']={'dictionary':f.filename,'entries':len(dictionary),'missing_count':len(missing),'missing_path':missing_path}
    d['work_language']='en'; d['sub_df']=None; d['suggestions']=None; d['depth']=None
    return jsonify({'ok':True,'files':_file_meta(d,eng),'stats':stats,'entries':len(dictionary),
                    'missing_count':len(missing),'missing_sample':sorted(missing)[:40]})

@app.route('/api/submissions/language',methods=['POST','OPTIONS'])
def submissions_language():
    body=request.get_json(force=True,silent=True) or {}
    sid,d=get_session(body.get('session_id'))
    if body.get('mode') not in ('ar','en'):return jsonify({'ok':False,'error':'Choose Arabic or English.'}),400
    d['work_language']=body['mode']
    return jsonify({'ok':True,'mode':d['work_language']})

@app.route('/api/submissions/translation-download',methods=['GET'])
def translation_download():
    if TR is None: return _tr_unavailable()
    sid=request.args.get('session_id'); _,d=get_session(sid); kind=request.args.get('kind','missing')
    if kind=='missing':
        path=(d.get('translation') or {}).get('missing_path')
        if not path or not os.path.exists(path):return jsonify({'ok':False,'error':'No translation has been run yet.'}),400
        return xlsx_response(path,'Missing_Arabic_Words.xlsx')
    label=request.args.get('file','')
    sf=next((x for x in d['sub_files'] if x['label']==label),None)
    if not sf or not sf.get('orig_path'):return jsonify({'ok':False,'error':'That file was not translated.'}),400
    return xlsx_response(sf['path'],Path(label).stem+'_translated'+Path(sf['path']).suffix)


# ---------------- Off-balance: the Outstanding Report and its two pivots ----------------

def _ob_unavailable():
    return jsonify({'ok':False,'error':'The Outstanding Report needs the latest bahrain_iraq_algorithm library in your project '
                    '(copy outstanding_report.py and engine.py into it, then restart the backend). Detail: %s'%OR_ERROR}),400

def _flat(f):
    return {'title':f['title'],'label_cols':f['label_cols'],'columns':f['columns'],'category':f.get('category'),
            'rows':[{'level':r['level'],'labels':r['labels'],'values':r['values'],'total':r['total']} for r in f['rows']]}

def _ob_result_payload(res,name,status):
    return {'ok':True,'source':name,'sheet':res['info']['sheet'],'header_row':res['info']['header_row'],'warnings':res['warnings'],
            'validation':[[k,v if isinstance(v,(int,float)) else str(v)] for k,v in res['validation']],
            'pivot1':_flat(res['pivot1']['flat']),'pivot2':_flat(res['pivot2']['flat']),'checks':res['checks'],'spec':res['spec'],'status':status}

def _ob_payload(d):
    """The ACTIVE pivots (automatic, saved definition, or the one the user approved) plus what the pivot builder needs."""
    ob=d['outstanding']
    p=_ob_result_payload(ob['result'],ob['name'],ob.get('status','auto'))
    p['saved']=bool(OR.load_definition()); p['sheets']=ob.get('sheets') or []; p['saved_error']=ob.get('saved_error')
    p['definition_path']=OR.definition_path()
    return p

@app.route('/api/outstanding/upload',methods=['POST','OPTIONS'])
def outstanding_upload():
    if OR is None: return _ob_unavailable()
    sid,d=get_session(request.form.get('session_id')); f=request.files.get('file')
    if not f:return jsonify({'ok':False,'error':'Upload the Outstanding Report workbook.'}),400
    name=Path(f.filename or 'outstanding.xlsx').name
    path=save_upload(f,d,'outstanding')
    if Path(name).suffix.lower() in ('.xls','.xlsb'):
        if TR is None: return _tr_unavailable()
        path=TR.to_xlsx(path,d['dir'])
    try:
        res=OR.process_outstanding(path,name,sheet=request.form.get('sheet') or None)
    except OR.OutstandingError as exc:
        return jsonify({'ok':False,'error':str(exc)}),400
    ob={'path':path,'name':name,'auto_result':res,'result':res,'status':'auto','draft':None,'saved_error':None}
    try: ob['sheets']=OR.list_sheets(path)
    except Exception: ob['sheets']=[]
    saved=OR.load_definition()
    if saved:                                  # a pivot definition saved in the library runs by itself on every new report
        try:
            ob['result']=OR.build_from_spec(path,saved,name,'saved'); ob['status']='saved'
        except OR.OutstandingError as exc:
            ob['saved_error']='The saved pivot definition could not be applied to this report (%s) - the automatic pivots are used instead.'%exc
    d['outstanding']=ob; d['depth']=None
    return jsonify(_ob_payload(d))

def _ob_need(d):
    if not d.get('outstanding'): return jsonify({'ok':False,'error':'Upload the Outstanding Report first.'}),400
    return None

@app.route('/api/outstanding/build',methods=['POST','OPTIONS'])
def outstanding_build():
    """Pivot builder: build pivots from the sheet / header row / columns the user picked. A DRAFT only - it is not used
    anywhere until it is approved."""
    if OR is None: return _ob_unavailable()
    body=request.get_json(force=True,silent=True) or {}
    sid,d=get_session(body.get('session_id'))
    bad=_ob_need(d)
    if bad: return bad
    ob=d['outstanding']
    try: draft=OR.build_from_spec(ob['path'],body.get('spec') or {},ob['name'],'draft')
    except OR.OutstandingError as exc: return jsonify({'ok':False,'error':str(exc)}),400
    ob['draft']=draft
    return jsonify({'ok':True,'draft':_ob_result_payload(draft,ob['name'],'draft')})

@app.route('/api/outstanding/approve',methods=['POST','OPTIONS'])
def outstanding_approve():
    if OR is None: return _ob_unavailable()
    body=request.get_json(force=True,silent=True) or {}
    sid,d=get_session(body.get('session_id'))
    bad=_ob_need(d)
    if bad: return bad
    ob=d['outstanding']
    if not ob.get('draft'): return jsonify({'ok':False,'error':'Build the pivots first, check the totals, then approve.'}),400
    ob['result']=ob['draft']; ob['result']['status']='approved'; ob['status']='approved'; ob['draft']=None; ob['saved_error']=None
    d['depth']=None
    return jsonify(_ob_payload(d))

@app.route('/api/outstanding/auto',methods=['POST','OPTIONS'])
def outstanding_auto():
    """Go back to the automatically created pivots."""
    if OR is None: return _ob_unavailable()
    body=request.get_json(force=True,silent=True) or {}
    sid,d=get_session(body.get('session_id'))
    bad=_ob_need(d)
    if bad: return bad
    ob=d['outstanding']; ob['result']=ob['auto_result']; ob['status']='auto'; ob['draft']=None; d['depth']=None
    return jsonify(_ob_payload(d))

@app.route('/api/outstanding/save-definition',methods=['POST','OPTIONS'])
def outstanding_save_definition():
    """Save the ACTIVE pivot definition in the library (pivot_definition.json) so it runs automatically on every new report."""
    if OR is None: return _ob_unavailable()
    body=request.get_json(force=True,silent=True) or {}
    sid,d=get_session(body.get('session_id'))
    bad=_ob_need(d)
    if bad: return bad
    spec=d['outstanding']['result']['spec']
    try: path=OR.save_definition(spec)
    except OSError as exc:
        return jsonify({'ok':False,'error':'The library folder is not writable from here (%s). Download the definition and copy it into the bahrain_iraq_algorithm folder as %s.'%(exc,OR.DEFINITION_FILE),
                        'definition':spec}),400
    return jsonify({'ok':True,'path':path,'saved':True})

@app.route('/api/outstanding/delete-definition',methods=['POST','OPTIONS'])
def outstanding_delete_definition():
    if OR is None: return _ob_unavailable()
    return jsonify({'ok':True,'deleted':OR.delete_definition(),'saved':bool(OR.load_definition())})

@app.route('/api/outstanding/definition',methods=['GET'])
def outstanding_definition():
    if OR is None: return _ob_unavailable()
    _,d=get_session(request.args.get('session_id'))
    bad=_ob_need(d)
    if bad: return bad
    spec=dict(d['outstanding']['result']['spec'])
    doc={'format':'bahrain-iraq-outstanding-pivots','version':1,'sheet':spec.get('sheet'),'header_row':spec.get('header_row'),
         'pivot1':spec['pivot1'],'pivot2':spec['pivot2']}
    return bytes_response(json.dumps(doc,ensure_ascii=False,indent=1).encode('utf-8'),OR.DEFINITION_FILE,'application/json')

@app.route('/api/outstanding/download',methods=['GET'])
def outstanding_download():
    if OR is None: return _ob_unavailable()
    sid=request.args.get('session_id'); _,d=get_session(sid)
    if not d.get('outstanding'):return jsonify({'ok':False,'error':'Upload the Outstanding Report first.'}),400
    path=os.path.join(d['dir'],'Outstanding_Report_IRAQ_Pivot_Output.xlsx')
    OR.create_output_workbook(d['outstanding']['result'],path)
    return xlsx_response(path,'Outstanding_Report_IRAQ_Pivot_Output.xlsx')
