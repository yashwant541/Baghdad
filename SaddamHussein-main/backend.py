import os, sys, json, uuid, tempfile, shutil
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

@app.route('/api/session',methods=['POST'])
def session():
    sid,_=get_session(); return jsonify({'ok':True,'session_id':sid})

@app.route('/api/inspect',methods=['POST'])
def inspect():
    sid,d=get_session(request.form.get('session_id')); f=request.files.get('file')
    if not f:return jsonify({'ok':False,'error':'No file uploaded'}),400
    path=save_upload(f,d,'inspect'); return jsonify({'ok':True,'session_id':sid,'file':Path(f.filename).name,'sheets':ReconEngine().inspect_workbook(path)})

@app.route('/api/tb',methods=['POST'])
def upload_tb():
    sid,d=get_session(request.form.get('session_id')); f=request.files.get('file')
    if not f:return jsonify({'ok':False,'error':'Upload a Trial Balance workbook.'}),400
    try:
        path=save_upload(f,d,'tb'); eng=ReconEngine(); tb,meta=eng.load_tb(path,request.form.get('sheet') or None,request.form.get('header_row') or None)
        pivot=eng.make_pivot(tb); tree=eng.build_tb_tree(tb)
        d.update(tb_path=path,tb=tb,pivot=pivot,tb_tree=tree)
        return jsonify({'ok':True,'session_id':sid,'meta':meta,
                        'kpis':{'rows':len(tb),'groups':int(tb.bs_mapping.nunique()),'currencies':int(tb.tran_ccy.nunique()),
                                'balance':float(tb.adjusted_balance.sum()),'unmapped':int((tb.bs_mapping=='').sum())},
                        'preview':json.loads(pivot.head(100).to_json(orient='records')),'columns':pivot.columns.tolist(),
                        'tree':tree})
    except Exception as e:return jsonify({'ok':False,'error':str(e)}),400

@app.route('/api/submissions/upload',methods=['POST'])
def submissions_upload():
    sid,d=get_session(request.form.get('session_id')); fs=request.files.getlist('files')
    if not fs:return jsonify({'ok':False,'error':'Upload at least one submission workbook.'}),400
    try:
        eng=ReconEngine(); d['sub_files']=[]; d['sub_inspect']={}
        out=[]
        for i,f in enumerate(fs[:3],1):
            path=save_upload(f,d,f'sub{i}'); label=Path(f.filename).name
            d['sub_files'].append({'label':label,'path':path})
            sheets=eng.inspect_workbook(path)
            d['sub_inspect'][label]=sheets
            out.append({'file':label,'sheets':sheets})
        return jsonify({'ok':True,'files':out})
    except Exception as e:return jsonify({'ok':False,'error':str(e)}),400

@app.route('/api/submissions/extract',methods=['POST'])
def submissions_extract():
    body=request.get_json(force=True); sid,d=get_session(body.get('session_id'))
    if not d['sub_files']:return jsonify({'ok':False,'error':'Upload submission workbooks first.'}),400
    selection=body.get('selection',{})
    try:
        eng=ReconEngine(); frames=[]
        for entry in d['sub_files']:
            label,path=entry['label'],entry['path']
            sheets=selection.get(label)
            if not sheets: continue
            frames.append(eng.extract_submission(path,label,sheets=sheets))
        import pandas as pd
        d['sub_df']=pd.concat(frames,ignore_index=True) if frames else pd.DataFrame()
        d['suggestions']=eng.auto_match(d['tb_tree'] or [], d['sub_df']) if d['tb_tree'] is not None else []
        preview=d['sub_df'].copy()
        if 'hierarchy' in preview.columns: preview=preview.drop(columns=['hierarchy'])
        return jsonify({'ok':True,'count':len(d['sub_df']),'preview':json.loads(preview.head(400).to_json(orient='records')),
                        'suggestions':d['suggestions']})
    except Exception as e:return jsonify({'ok':False,'error':str(e)}),400

@app.route('/api/reconcile',methods=['POST'])
def reconcile():
    body=request.get_json(force=True); sid,d=get_session(body.get('session_id'))
    if d['tb'] is None or d['sub_df'] is None:return jsonify({'ok':False,'error':'Process the Trial Balance and submissions first.'}),400
    rules=body.get('rules',[])
    eng=ReconEngine(body.get('tolerance_abs',1),body.get('tolerance_pct',0.0001))
    d['rules']=rules
    d['recon']=eng.reconcile(d['tb'],d['sub_df'],rules)
    recon_rows=json.loads(d['recon'].to_json(orient='records'))
    lineage=eng.build_lineage(d['tb_tree'] or [], recon_rows)
    d['lineage']=lineage
    return jsonify({'ok':True,'results':recon_rows,'lineage':lineage})

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
    eng.export(d['tb'],d['pivot'],d['sub_df'],d['recon'],d.get('suggestions'),flat_lineage,path)
    return send_file(path,as_attachment=True,download_name='Bahrain_Iraq_Reconciliation_Output.xlsx')
