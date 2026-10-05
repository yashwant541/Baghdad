import os, sys, json, uuid, tempfile, traceback
from pathlib import Path
from flask import request, jsonify, send_file

HERE=os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
if HERE not in sys.path: sys.path.insert(0,HERE)
from bahrain_iraq_algorithm import ReconEngine

SESSIONS={}

def get_session(sid=None):
    sid=sid or str(uuid.uuid4())
    d=SESSIONS.setdefault(sid,{'dir':tempfile.mkdtemp(prefix='bahrain_iraq_'),'tb_path':None,'tb':None,'pivot':None,
                               'tb_tree':None,'sub_files':[],'sub_inspect':{},'sub_df':None,'suggestions':None,
                               'rules':[],'recon':None,'lineage':None})
    return sid,d

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
    return jsonify({'ok':True,'session_id':sid,'meta':meta,
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
    for i,f in enumerate(fs[:8],1):
        path=save_upload(f,d,f'sub{i}'); label=Path(f.filename).name
        d['sub_files'].append({'label':label,'path':path})
        sheets=eng.inspect_workbook(path)
        d['sub_inspect'][label]=sheets
        out.append({'file':label,'sheets':sheets})
    return jsonify({'ok':True,'files':out})

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
    default_matches=[]; coverage=[]
    if d['tb'] is not None and len(d['sub_df']):
        templates=eng.load_default_mapping()
        for tpl in templates:
            resolved,warnings=eng.resolve_rule_template(tpl,d['tb'],d['sub_df'])
            tb_n=len(resolved.get('tb_components') or []); sub_n=len(resolved.get('components') or [])
            fulfilled=tb_n>0 and sub_n>0
            entry={'label':resolved.get('label'),'currency':resolved.get('currency'),'rule_type':resolved.get('rule_type'),
                   'fulfilled':fulfilled,'tb_found':tb_n>0,'sub_found':sub_n>0,
                   'tb_accounts':tb_n,'sub_lines':sub_n,'warnings':warnings,'reconciled_status':None}
            if fulfilled:
                resolved['source']='DEFAULT'
                default_matches.append({'resolved':resolved,'warnings':warnings})
                entry['tb_amount']=sum(c['amount']*c.get('sign',1) for c in resolved['tb_components'])
                entry['submission_amount']=sum(c['amount']*c.get('sign',1)*c.get('multiplier',1) for c in resolved['components'])
                entry['difference']=entry['tb_amount']-entry['submission_amount']
            coverage.append(entry)
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
                    'mapping_coverage':coverage,'mapping_coverage_summary':summary})

@app.route('/api/submissions/search-amount',methods=['POST','OPTIONS'])
def submissions_search_amount():
    body=request.get_json(force=True,silent=True) or {}
    sid,d=get_session(body.get('session_id'))
    if d.get('sub_df') is None or not len(d['sub_df']):
        return jsonify({'ok':False,'error':'Process submissions first.'}),400
    amount=body.get('amount')
    if amount is None:return jsonify({'ok':False,'error':'No amount given.'}),400
    eng=ReconEngine()
    results=eng.search_amount(d['sub_df'],amount,body.get('tolerance_abs',1),body.get('tolerance_pct',0.0001))
    return jsonify({'ok':True,'results':results[:200],'total_found':len(results)})

@app.route('/api/submissions/export',methods=['GET'])
def submissions_export():
    sid=request.args.get('session_id'); _,d=get_session(sid)
    if d.get('sub_df') is None or not len(d['sub_df']):return jsonify({'ok':False,'error':'No submission lines extracted yet.'}),400
    eng=ReconEngine()
    path=os.path.join(d['dir'],'Submissions_Simplified.xlsx')
    eng.export_submissions(d['sub_df'],path)
    return send_file(path,as_attachment=True,download_name='Submissions_Simplified.xlsx')

@app.route('/api/rules/resolve',methods=['POST','OPTIONS'])
def rules_resolve():
    body=request.get_json(force=True,silent=True) or {}
    sid,d=get_session(body.get('session_id'))
    if d.get('tb') is None or d.get('sub_df') is None:
        return jsonify({'ok':False,'error':'Upload a Trial Balance and process submissions before importing a mapping.'}),400
    templates=body.get('templates',[])
    if not templates:return jsonify({'ok':False,'error':'The uploaded mapping file has no matches in it.'}),400
    eng=ReconEngine(); out=[]
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
    return send_file(path,as_attachment=True,download_name='Bahrain_Iraq_Reconciliation_Output.xlsx')
