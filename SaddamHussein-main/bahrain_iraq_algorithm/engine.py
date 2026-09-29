import io, os, re, json, difflib, itertools, tempfile
from datetime import datetime
from pathlib import Path
from collections import defaultdict
import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter, column_index_from_string

ALIASES = {
 "account": ["account", "gl account", "account no", "account number"],
 "account_desc": ["account desc", "account description", "gl description"],
 "tran_ccy": ["tran ccy", "transaction currency", "currency", "ccy"],
 "bs_mapping": ["bs mapping", "balance sheet mapping", "bs group", "mapping"],
 "adjusted_balance": ["adjusted balance", "adjusted bal", "ytd adjusted balance"],
 "base_amt": ["base amt", "base amount", "local currency amount", "lc amount"],
 "adjustment": ["adjustment", "adjustments"],
 "category": ["cat", "category"],
 "pl_mapping": ["p&l mapping", "p/l mapping", "pl mapping"]
}

CURRENCY_CODES = ["USD","IQD","BHD","EUR","GBP","AED","SAR","KWD","QAR","OMR","JOD","CHF","JPY","CNY"]
CURRENCY_PATTERN = re.compile(r"\b(" + "|".join(CURRENCY_CODES) + r")\b", re.I)
LOCAL_CCY_PATTERN = re.compile(r"\blocal\s*currency\b|\blcy\b", re.I)
# many real-world templates (this Central Bank of Iraq one included) spell the
# currency out by name rather than by ISO code — "Accounts in Iraqi Dinars",
# "Foreign Currency Accounts Converted to Iraqi Dinars" — so the ISO-code
# regex alone misses every column. Checked in order: qualified/unambiguous
# names first, bare "Dollars"/"Dirhams" (common shorthand for USD/AED) last.
CURRENCY_NAME_PATTERNS = [
    (re.compile(r"\biraqi\s+dinars?\b", re.I), "IQD"),
    (re.compile(r"\bbahraini\s+dinars?\b", re.I), "BHD"),
    (re.compile(r"\bkuwaiti\s+dinars?\b", re.I), "KWD"),
    (re.compile(r"\bjordanian\s+dinars?\b", re.I), "JOD"),
    (re.compile(r"\bu\.?s\.?\s*dollars?\b", re.I), "USD"),
    (re.compile(r"\bsaudi\s+riyals?\b", re.I), "SAR"),
    (re.compile(r"\bqatari\s+riyals?\b", re.I), "QAR"),
    (re.compile(r"\bomani\s+rials?\b", re.I), "OMR"),
    (re.compile(r"\b(emirati|uae)\s+dirhams?\b", re.I), "AED"),
    (re.compile(r"\beuros?\b", re.I), "EUR"),
    (re.compile(r"\bpound\s*sterling\b|\bsterling\b", re.I), "GBP"),
    (re.compile(r"\bdollars?\b", re.I), "USD"),
    (re.compile(r"\bdirhams?\b", re.I), "AED"),
]
def match_currency_name(text):
    for pat, code in CURRENCY_NAME_PATTERNS:
        if pat.search(text): return code
    return None

TOTAL_PATTERN = re.compile(r"\b(grand\s+total|sub\s*-?\s*total|total)\b", re.I)
# a column-index / cross-reference row some regulatory templates print under
# the real header — e.g. "2=(3+4+5+6)" — reads as plausible-looking small
# integers per column and must never be mistaken for a real data line.
FORMULA_REF_PATTERN = re.compile(r"\d+\s*=\s*\([^)]*[+\-][^)]*\)")
TOTAL_PREFIX = re.compile(r"^(grand\s+total\s+of\s+|grand\s+total\s+|sub[\s-]?total\s+of\s+|sub[\s-]?total\s+|total\s+of\s+|total\s+)", re.I)
SCALE_THOUSANDS = re.compile(r"\(?'?000\b\)?|in\s+thousands?\b|\bthousands?\b", re.I)
SCALE_MILLIONS = re.compile(r"\(?'?mn'?\)?|in\s+millions?\b|\bmillions?\b", re.I)
HIERARCHY_DELIMS = [" > ", " / ", " - ", " – ", " :: ", " | "]
STOPWORDS = {"and","the","of","for","in","on","to","a","an","with","&"}
# common reporting-scale multipliers to try numerically when a header's text
# gives no '000/million clue at all (e.g. a bare "USD" column that's quietly
# already divided by a thousand) — 1.0 is handled separately as the default.
SCALE_CANDIDATES = (1000.0, 1000000.0, 0.001, 0.000001)

def strip_total_prefix(text):
    t = str(text or "").strip()
    s = TOTAL_PREFIX.sub("", t).strip()
    return s or t

def forward_fill_headers(headers):
    """Many real templates visually group several sub-columns under one
    label ("Accounts in Iraqi Dinars" over Residents/Non-Residents) without
    an actual Excel cell merge — the label only lives in the leftmost cell,
    the rest of the group reads back blank. Carry it rightward so every
    column in the group still classifies correctly."""
    out=[]; last=''
    for h in headers:
        h=str(h or '').strip()
        if h: last=h
        out.append(last)
    return out

def detect_scale(text):
    t = str(text or "")
    if SCALE_THOUSANDS.search(t): return 1000.0
    if SCALE_MILLIONS.search(t): return 1000000.0
    return 1.0

def norm(v):
    return re.sub(r"[^a-z0-9]+", " ", str(v or "").strip().lower()).strip()

def norm_tokens(v):
    t = re.sub(r"[^a-z0-9\s]", " ", str(v or "").lower())
    return [w for w in t.split() if w and w not in STOPWORDS]

def text_similarity(a, b):
    ta, tb = set(norm_tokens(a)), set(norm_tokens(b))
    seq = difflib.SequenceMatcher(None, str(a or "").lower(), str(b or "").lower()).ratio()
    if not ta or not tb:
        return seq
    jaccard = len(ta & tb) / len(ta | tb) if (ta | tb) else 0.0
    return max(jaccard, seq)

def classify_currency_header(text):
    t = str(text or "").strip()
    if not t:
        return None
    is_total = bool(TOTAL_PATTERN.search(t))
    m = CURRENCY_PATTERN.search(t)
    code = m.group(1).upper() if m else None
    if not code:
        code = match_currency_name(t)
    if not code and LOCAL_CCY_PATTERN.search(t):
        code = "LCY"
    if is_total and code:
        return f"TOTAL_{code}"
    if is_total:
        return "TOTAL"
    return code

def split_hierarchy(text):
    t = str(text or "").strip()
    if not t:
        return []
    for d in HIERARCHY_DELIMS:
        if d in t:
            return [p.strip() for p in t.split(d) if p.strip()]
    return [t]

def json_value(v):
    if pd.isna(v): return None
    if isinstance(v, (np.integer,)): return int(v)
    if isinstance(v, (np.floating,)): return float(v)
    if isinstance(v, (pd.Timestamp, datetime)): return v.isoformat()
    return str(v) if not isinstance(v,(str,int,float,bool)) else v

class ReconEngine:
    def __init__(self, tolerance_abs=1.0, tolerance_pct=0.0001):
        self.tolerance_abs=float(tolerance_abs); self.tolerance_pct=float(tolerance_pct)

    # ---------- workbook inspection ----------

    def _read_raw(self, path, sheet):
        return pd.read_excel(path, sheet_name=sheet, header=None, dtype=object, engine='openpyxl')

    def _preamble_scale(self, raw, header_row):
        """A scale note ("Amounts in Thousand Iraqi Dinars") very often sits
        in the report's title block above the column headers, not inside any
        header cell itself. Scan that preamble region once per sheet and use
        it as the default for any currency column whose own header text
        didn't state a scale."""
        if header_row is None: return 1.0
        for i in range(header_row):
            for v in raw.iloc[i].tolist():
                s=detect_scale(v)
                if s!=1.0: return s
        return 1.0

    def inspect_workbook(self, path):
        xls=pd.ExcelFile(path, engine='openpyxl'); out=[]
        for sheet in xls.sheet_names:
            raw=self._read_raw(path,sheet)
            header,score=self.detect_header(raw)
            headers=[] if header is None else [str(x).strip() if pd.notna(x) else '' for x in raw.iloc[header].tolist()]
            preamble_scale=self._preamble_scale(raw,header)
            filled_headers=forward_fill_headers(headers)
            currency_columns={}
            for j,h in enumerate(filled_headers):
                code=classify_currency_header(h)
                if code:
                    scale=detect_scale(h)
                    if scale==1.0: scale=preamble_scale
                    currency_columns[get_column_letter(j+1)]={'currency':code,'scale':scale}
            out.append({'sheet':sheet,'rows':int(raw.shape[0]),'columns':int(raw.shape[1]),
                        'header_row':None if header is None else int(header+1),'score':score,
                        'headers':headers[:60],'currency_columns':currency_columns})
        return out

    def detect_header(self, raw, max_rows=80):
        best=(None,-1)
        vocabulary=set(sum(ALIASES.values(),[])) | {'amount','description','line item','particulars','total'}
        for i in range(min(max_rows,len(raw))):
            cells=[norm(x) for x in raw.iloc[i].tolist() if pd.notna(x)]
            score=sum(3 for c in cells if c in vocabulary)+sum(1 for c in cells if any(k in c for k in ['account','amount','balance','mapping','description']))
            if len(set(cells))>=3 and score>best[1]: best=(i,score)
        return best if best[1]>=3 else (None,best[1])

    def _canonicalize(self, df):
        cols={norm(c):c for c in df.columns}; rename={}
        for target, aliases in ALIASES.items():
            for a in aliases:
                if norm(a) in cols: rename[cols[norm(a)]]=target; break
        return df.rename(columns=rename), rename

    # ---------- Trial Balance ----------

    def load_tb(self, path, sheet=None, header_row=None):
        inspections=self.inspect_workbook(path)
        if not inspections: raise ValueError('No worksheets found.')
        pick=next((x for x in inspections if x['sheet']==sheet),None) if sheet else max(inspections,key=lambda x:x['score'])
        if not pick: raise ValueError('Selected worksheet was not found.')
        hr=int(header_row)-1 if header_row else pick['header_row']-1 if pick['header_row'] else None
        if hr is None: raise ValueError('Could not detect the Trial Balance header row.')
        df=pd.read_excel(path,sheet_name=pick['sheet'],header=hr,dtype=object,engine='openpyxl')
        df.columns=[str(c).strip() for c in df.columns]
        df,rename=self._canonicalize(df); required=['account','account_desc','tran_ccy','bs_mapping','adjusted_balance']
        missing=[x for x in required if x not in df.columns]
        if missing: raise ValueError('Missing required Trial Balance fields: '+', '.join(missing))
        for c in ['account','account_desc','tran_ccy','bs_mapping']:
            df[c]=df[c].fillna('').astype(str).str.strip()
        df['adjusted_balance']=pd.to_numeric(df['adjusted_balance'],errors='coerce').fillna(0.0)
        for c in ['base_amt','adjustment']:
            if c in df: df[c]=pd.to_numeric(df[c],errors='coerce').fillna(0.0)
        df=df.loc[~((df['account']=='') & (df['account_desc']=='') & (df['adjusted_balance']==0))].copy()
        df['_source_row']=np.arange(hr+2,hr+2+len(df)); df['_source_sheet']=pick['sheet']
        return df, {'sheet':pick['sheet'],'header_row':hr+1,'rows':len(df),'column_map':{str(k):v for k,v in rename.items()}}

    def make_pivot(self, tb):
        p=pd.pivot_table(tb,index=['bs_mapping','account','account_desc'],columns='tran_ccy',values='adjusted_balance',aggfunc='sum',fill_value=0,margins=True,margins_name='Grand Total').reset_index()
        p.columns=[str(c) for c in p.columns]
        return p

    def build_tb_tree(self, tb):
        tb=tb.copy()
        tb['_hierarchy']=tb['bs_mapping'].apply(split_hierarchy)
        tb['_top']=tb['_hierarchy'].apply(lambda h: h[0] if h else '(Unmapped)')
        tb['_sub']=tb['_hierarchy'].apply(lambda h: h[1] if len(h)>1 else '')
        tree=[]
        for top, gdf in tb.groupby('_top', sort=False):
            node={'name':top,'level':'group','currency_totals':self._currency_totals(gdf),'children':[]}
            subs=gdf[gdf['_sub']!='']
            if len(subs):
                for sub,sdf in subs.groupby('_sub', sort=False):
                    node['children'].append({'name':sub,'level':'subgroup','currency_totals':self._currency_totals(sdf),'accounts':self._account_rows(sdf),'children':[]})
                leftover=gdf[gdf['_sub']=='']
                if len(leftover):
                    node['children'].append({'name':'(Direct)','level':'subgroup','currency_totals':self._currency_totals(leftover),'accounts':self._account_rows(leftover),'children':[]})
            else:
                node['accounts']=self._account_rows(gdf)
            tree.append(node)
        return tree

    def _currency_totals(self, df):
        out={str(k):float(v) for k,v in df.groupby('tran_ccy')['adjusted_balance'].sum().to_dict().items()}
        out['TOTAL']=float(df['adjusted_balance'].sum())
        return out

    def _account_rows(self, df):
        return [{'account':r.account,'account_desc':r.account_desc,'currency':r.tran_ccy,'amount':float(r.adjusted_balance),'bs_mapping':r.bs_mapping} for r in df.itertuples()]

    # ---------- Submissions ----------

    def _to_number(self, v):
        if isinstance(v,(int,float,np.number)) and not pd.isna(v): return float(v)
        if isinstance(v,str):
            s=v.replace(',','').replace('(','-').replace(')','').replace('%','').strip()
            if re.fullmatch(r'-?\d+(\.\d+)?', s) and s not in ('','-'): return float(s)
        return None

    def extract_submission(self, path, file_label=None, sheets=None):
        file_label = file_label or Path(path).name
        inspections = self.inspect_workbook(path)
        insp_by_name = {x['sheet']: x for x in inspections}
        if sheets:
            targets=[]
            for s in sheets:
                name = s['sheet'] if isinstance(s, dict) else s
                info = dict(insp_by_name.get(name) or {'sheet':name,'header_row':None,'currency_columns':{}})
                override = s.get('header_row') if isinstance(s, dict) else None
                if override:
                    info['header_row']=int(override)
                    raw=self._read_raw(path,name); hr=int(override)-1
                    headers=[str(x).strip() if pd.notna(x) else '' for x in raw.iloc[hr].tolist()] if hr < len(raw) else []
                    preamble_scale=self._preamble_scale(raw,hr)
                    filled_headers=forward_fill_headers(headers)
                    info['currency_columns']={}
                    for j,h in enumerate(filled_headers):
                        code=classify_currency_header(h)
                        if code:
                            scale=detect_scale(h)
                            if scale==1.0: scale=preamble_scale
                            info['currency_columns'][get_column_letter(j+1)]={'currency':code,'scale':scale}
                targets.append(info)
        else:
            targets = inspections

        wb=load_workbook(path, data_only=True)
        rows=[]
        for info in targets:
            sheet_name=info['sheet']
            if sheet_name not in wb.sheetnames: continue
            raw=self._read_raw(path, sheet_name); ws=wb[sheet_name]
            header_row = info['header_row']-1 if info.get('header_row') else None
            currency_cols={column_index_from_string(k)-1:v for k,v in (info.get('currency_columns') or {}).items()}
            has_any_currency=bool(currency_cols)
            stack=[]
            for r in range(len(raw)):
                if header_row is not None and r<=header_row: continue
                vals=raw.iloc[r].tolist()
                # a column-index / cross-reference row (e.g. "2=(3+4+5+6)")
                # is template scaffolding, not data — never mistake its small
                # integers per column for a real financial line.
                if any(pd.notna(v) and FORMULA_REF_PATTERN.search(str(v)) for v in vals):
                    continue
                # the label is the longest NON-numeric cell in the row, not
                # simply the first non-empty one — many real templates put a
                # bare serial number ("1") or letter ("A") in the leftmost
                # column, with the actual account description several
                # columns over. Falls back to "first non-empty" only if every
                # non-empty cell in the row happens to be numeric-typed.
                label_col,label,best_len=None,'',-1
                for j,v in enumerate(vals):
                    if not pd.notna(v): continue
                    s=str(v).strip()
                    if not s or self._to_number(v) is not None: continue
                    if len(s)>best_len: label_col,label,best_len=j,s,len(s)
                if label_col is None:
                    for j,v in enumerate(vals):
                        if pd.notna(v) and str(v).strip():
                            label_col,label=j,str(v).strip(); break
                if label_col is None: continue
                numeric_cells={}
                for j,v in enumerate(vals):
                    if j==label_col: continue
                    num=self._to_number(v)
                    if num is not None: numeric_cells[j]=num
                cell=ws.cell(row=r+1, column=label_col+1)
                indent=0
                try: indent=int(cell.alignment.indent or 0)
                except Exception: indent=0
                bold=bool(cell.font and cell.font.bold)
                is_total=bool(TOTAL_PATTERN.search(label))
                is_header_row = not numeric_cells
                if is_header_row:
                    if bold or label.isupper() or indent>0:
                        level = indent if indent else 0
                        while stack and stack[-1][0] >= level: stack.pop()
                        stack.append((level, label))
                    continue
                hierarchy_path=[name for _,name in stack]
                level = indent if indent else len(stack)
                for j, amount in numeric_cells.items():
                    colinfo=currency_cols.get(j)
                    if not colinfo:
                        if has_any_currency: continue
                        ccy,scale='UNSPECIFIED',1.0
                    else:
                        ccy,scale=colinfo.get('currency','UNSPECIFIED'),float(colinfo.get('scale',1.0))
                    rows.append({
                        'submission_file':file_label,'sheet':sheet_name,'row_number':r+1,
                        'source_cell':f'{get_column_letter(j+1)}{r+1}','label_cell':f'{get_column_letter(label_col+1)}{r+1}',
                        'line_description':label,'hierarchy':hierarchy_path,
                        'hierarchy_path':' / '.join(hierarchy_path) if hierarchy_path else '',
                        'level':level,'is_total':is_total,'currency':ccy,'scale':scale,
                        'raw_amount':amount,'normalized_amount':amount*scale,'unit_multiplier':scale,'sign':1.0
                    })
        return pd.DataFrame(rows)

    # ---------- Auto-matching ----------

    def find_submission_subset(self, items, target, tolerance_abs, tolerance_pct, max_combo=4, max_items=10, scale=1.0):
        """Bounded subset-sum: which specific submission lines (out of a
        already section-matched pool) actually club up to a TB figure, when
        blindly summing every line in the pool doesn't reconcile and there's
        no sheet-provided Total row to trust instead. Deliberately narrow in
        scope — only ever called on items already confined to one matched TB
        group/sub-group and one matched submission section — so a numeric
        coincidence across unrelated parts of the file can't be suggested.
        `scale` multiplies each line's amount first, so the same search also
        covers "the right lines, but also reported at the wrong scale"."""
        n=len(items)
        if n==0 or n>max_items: return None
        tol=max(tolerance_abs, abs(target)*tolerance_pct)
        best=None
        for size in range(1, min(max_combo,n)+1):
            for combo in itertools.combinations(range(n), size):
                s=sum(items[i]['normalized_amount'] for i in combo)*scale
                diff=abs(target-s)
                if diff<=tol:
                    coverage=size/n
                    rank=(diff,-coverage)
                    if best is None or rank<best[0]:
                        best=(rank,combo)
        return best[1] if best else None

    def best_scale_fit(self, raw_value, target, tolerance_abs, tolerance_pct, scales=SCALE_CANDIDATES):
        """Beyond whatever scale a column header's text implied at extraction
        time, check whether a common reporting-scale multiplier (thousands,
        millions, or the reverse) — together with a possible sign flip —
        would reconcile a raw value against the TB target on its own. Used
        both by the auto-matcher and, client-side in JS, mirrored by the
        manual match builder's live preview. Returns (scale, sign, diff) for
        the closest fit within tolerance, or None."""
        tol=max(tolerance_abs, abs(target)*tolerance_pct)
        best=None
        for scale in scales:
            for sign in (1.0,-1.0):
                s=raw_value*scale*sign
                diff=abs(target-s)
                if diff<=tol and (best is None or diff<best[2]):
                    best=(scale,sign,diff)
        return best

    def auto_match(self, tb_tree, submission_df, min_group_score=0.35, min_sub_score=0.3):
        if submission_df is None or submission_df.empty: return []
        recs=submission_df.to_dict('records')
        sub_top=defaultdict(list)
        for rec in recs:
            hier = rec.get('hierarchy') or []
            top = hier[0] if hier else rec['line_description']
            sub_top[top].append(rec)

        suggestions=[]
        for group in tb_tree:
            gname=group['name']
            best_top,best_score=None,0.0
            for top in sub_top:
                score=text_similarity(gname, top)
                if score>best_score: best_top,best_score=top,score
            if best_score < min_group_score or best_top is None: continue
            top_rows=sub_top[best_top]
            children = group.get('children') or [{'name':'(Direct)','level':'subgroup','currency_totals':group.get('currency_totals',{}),'accounts':group.get('accounts',[])}]

            # Partition every submission line under this section across the
            # group's sub-groups by best name fit, so a flat list of lines
            # with no bold sub-headers and no "Total X" rows (a very common
            # "imperfect" layout) still lands every relevant sibling line in
            # the same pool — not just whichever single line happens to score
            # highest. A dedicated "Total <sub-group>" row is authoritative
            # for that sub-group specifically and wins its pool outright.
            if len(children)==1 and children[0]['name']=='(Direct)':
                child_pools={'(Direct)':top_rows}
                child_match_score={'(Direct)':best_score}
                child_match_label={'(Direct)':best_top}
            else:
                child_pools={c['name']:[] for c in children}
                child_match_score={c['name']:0.0 for c in children}
                child_match_label={c['name']:None for c in children}
                for r in top_rows:
                    # a grand-total row for the WHOLE section ("TOTAL ASSETS")
                    # must never be claimed by one sub-group's pool — short,
                    # generic labels like that can otherwise out-score a real
                    # sub-group name on pure character similarity and silently
                    # double-count the section total inside one sub-group.
                    if r['is_total'] and (text_similarity(strip_total_prefix(r['line_description']),gname)>=0.6
                                           or text_similarity(strip_total_prefix(r['line_description']),best_top)>=0.6):
                        continue
                    hier=r.get('hierarchy') or []
                    candidate_label = hier[1] if len(hier)>1 else strip_total_prefix(r['line_description'])
                    best_child,best_child_score=None,0.0
                    for c in children:
                        if c['name']=='(Direct)': continue
                        score=text_similarity(c['name'],candidate_label)
                        if score>best_child_score: best_child,best_child_score=c['name'],score
                    if best_child is not None and best_child_score>=min_sub_score:
                        child_pools[best_child].append(r)
                        if best_child_score>child_match_score[best_child]:
                            child_match_score[best_child]=best_child_score
                            child_match_label[best_child]=candidate_label

            for child in children:
                cname=child['name']
                pool=child_pools.get(cname) or []
                if not pool: continue
                best_sub=child_match_label.get(cname)
                best_sub_score=child_match_score.get(cname,0.0)
                for ccy, tb_amt in (child.get('currency_totals') or {}).items():
                    if ccy=='TOTAL':
                        matches=[r for r in pool if r['currency']=='TOTAL' or str(r['currency']).startswith('TOTAL_')]
                    else:
                        matches=[r for r in pool if r['currency']==ccy]
                    if not matches: continue
                    totals=[r for r in matches if r['is_total']]
                    chosen = totals if totals else matches
                    raw_sum=sum(r['normalized_amount'] for r in chosen)
                    resolved_sign = -1.0 if abs(tb_amt-(-raw_sum)) < abs(tb_amt-raw_sum) else 1.0
                    sub_amt=raw_sum*resolved_sign
                    diff=tb_amt-sub_amt
                    label = gname if cname in ('(Direct)','') else f'{gname} / {cname}'
                    suggestions.append({
                        'bs_mapping':label,'group':gname,'subgroup':None if cname=='(Direct)' else cname,
                        'currency':ccy,'tb_amount':tb_amt,'suggested_submission_amount':sub_amt,'difference':diff,
                        'sign_applied':resolved_sign,
                        'confidence':round(min(best_score, best_sub_score if best_sub_score else best_score),2),
                        'match_basis':{'section_match':best_top,'subgroup_match':best_sub,'used_total_rows':bool(totals)},
                        'rule_type':'AUTO_TOTAL' if ccy=='TOTAL' else 'AUTO_CURRENCY',
                        'components':[{'submission_file':r['submission_file'],'sheet':r['sheet'],'row_number':r['row_number'],
                                       'source_cell':r['source_cell'],'line_description':r['line_description'],'currency':r['currency'],
                                       'amount':r['normalized_amount'],'multiplier':1,'sign':resolved_sign} for r in chosen]
                    })

                    # the blanket "sum every line in this section" suggestion above
                    # didn't reconcile and there was no Total row to fall back on —
                    # try to find which specific lines actually club up to the TB
                    # figure (e.g. one stray unrelated line in the same section is
                    # often the entire reason the blanket sum was off).
                    out_of_tolerance = abs(diff) > max(self.tolerance_abs, abs(tb_amt)*self.tolerance_pct)

                    # a bare column header (no '000 / million wording at all)
                    # can still hide a uniform reporting-scale factor — check
                    # numerically, independent of whatever scale the header
                    # text implied back at extraction time.
                    if out_of_tolerance:
                        fit=self.best_scale_fit(raw_sum, tb_amt, self.tolerance_abs, self.tolerance_pct)
                        if fit:
                            scale_f,sign_f,_=fit
                            sub_amt3=raw_sum*scale_f*sign_f
                            diff3=tb_amt-sub_amt3
                            suggestions.append({
                                'bs_mapping':label,'group':gname,'subgroup':None if cname=='(Direct)' else cname,
                                'currency':ccy,'tb_amount':tb_amt,'suggested_submission_amount':sub_amt3,'difference':diff3,
                                'sign_applied':sign_f,'scale_applied':scale_f,
                                'confidence':round(min(best_score,best_sub_score if best_sub_score else best_score)*0.85,2),
                                'match_basis':{'section_match':best_top,'subgroup_match':best_sub,'used_total_rows':bool(totals),
                                                'scale_detected':scale_f},
                                'rule_type':'AUTO_SCALE_ADJUST',
                                'components':[{'submission_file':r['submission_file'],'sheet':r['sheet'],'row_number':r['row_number'],
                                               'source_cell':r['source_cell'],'line_description':r['line_description'],'currency':r['currency'],
                                               'amount':r['normalized_amount'],'multiplier':scale_f,'sign':sign_f} for r in chosen]
                            })

                    if ccy!='TOTAL' and not totals and out_of_tolerance and len(matches)>1:
                        for scale_try in (1.0,)+SCALE_CANDIDATES:
                            combo_idx=self.find_submission_subset(matches, tb_amt*resolved_sign, self.tolerance_abs, self.tolerance_pct, scale=scale_try)
                            if combo_idx and len(combo_idx)<len(matches):
                                chosen2=[matches[i] for i in combo_idx]
                                sub_amt2=sum(r['normalized_amount'] for r in chosen2)*scale_try*resolved_sign
                                diff2=tb_amt-sub_amt2
                                basis={'section_match':best_top,'subgroup_match':best_sub,'used_total_rows':False,
                                       'auto_clubbed':True,'excluded_lines':len(matches)-len(chosen2)}
                                if scale_try!=1.0: basis['scale_detected']=scale_try
                                suggestions.append({
                                    'bs_mapping':label,'group':gname,'subgroup':None if cname=='(Direct)' else cname,
                                    'currency':ccy,'tb_amount':tb_amt,'suggested_submission_amount':sub_amt2,'difference':diff2,
                                    'sign_applied':resolved_sign,'scale_applied':scale_try,
                                    'confidence':round(min(best_score,best_sub_score if best_sub_score else best_score)*(0.9 if scale_try==1.0 else 0.8),2),
                                    'match_basis':basis,
                                    'rule_type':'AUTO_CLUB_SUBSET' if scale_try==1.0 else 'AUTO_CLUB_SUBSET_SCALED',
                                    'components':[{'submission_file':r['submission_file'],'sheet':r['sheet'],'row_number':r['row_number'],
                                                   'source_cell':r['source_cell'],'line_description':r['line_description'],'currency':r['currency'],
                                                   'amount':r['normalized_amount'],'multiplier':scale_try,'sign':resolved_sign} for r in chosen2]
                                })
                                break  # first (smallest-scale, preferring unscaled) fit wins
        suggestions.sort(key=lambda s: (-s['confidence'], s['bs_mapping']))
        return suggestions

    # ---------- Reconciliation ----------

    def reconcile(self, tb, submissions, rules):
        tb=tb.copy()
        tb['_hier_key']=tb['bs_mapping'].apply(lambda x: ' / '.join(split_hierarchy(x)))
        grp_raw_total=tb.groupby('bs_mapping')['adjusted_balance'].sum().to_dict()
        grp_raw_ccy=tb.groupby(['bs_mapping','tran_ccy'])['adjusted_balance'].sum().to_dict()
        grp_hier_total=tb.groupby('_hier_key')['adjusted_balance'].sum().to_dict()
        grp_hier_ccy=tb.groupby(['_hier_key','tran_ccy'])['adjusted_balance'].sum().to_dict()
        out=[]
        for rule in rules:
            bs=(rule.get('bs_mapping') or rule.get('label') or '').strip()
            ccy=rule.get('currency','TOTAL') or 'TOTAL'
            comps=rule.get('components',[]); sub_amt=0.0; evidence=[]
            for comp in comps:
                f,s,row=comp.get('submission_file'),comp.get('sheet'),comp.get('row_number')
                if row in (None,''): continue
                cell=comp.get('source_cell'); comp_ccy=comp.get('currency')
                q=(submissions.submission_file==f)&(submissions.sheet==s)&(submissions.row_number==int(row))
                if cell: q=q&(submissions.source_cell==cell)
                elif comp_ccy: q=q&(submissions.currency==comp_ccy)
                hit=submissions[q]
                if len(hit):
                    r0=hit.iloc[0]
                    amount=float(r0.normalized_amount)*float(comp.get('multiplier',1))*float(comp.get('sign',1))
                    sub_amt+=amount
                    evidence.append({'file':f,'sheet':s,'row':int(row),'description':r0.line_description,'amount':amount,'currency':str(r0.get('currency',''))})

            # the TB side of a rule is normally a single BS-Mapping group/currency
            # looked up from the Trial Balance totals (auto-suggestions work this
            # way). A manually built or imported match instead names its own list
            # of clubbed-up TB accounts/groups with real amounts already attached
            # (see build_tb_components / resolve_rule_template) — sum those directly
            # and keep the same per-item evidence trail the submission side has.
            tb_components=rule.get('tb_components')
            if tb_components:
                tb_amt=0.0; tb_evidence=[]
                for c in tb_components:
                    amt=float(c.get('amount',0))*float(c.get('sign',1))
                    tb_amt+=amt
                    tb_evidence.append({'account':c.get('account'),'description':c.get('account_desc'),
                                         'bs_mapping':c.get('bs_mapping'),'currency':c.get('currency'),'amount':amt})
            else:
                tb_evidence=None
                if ccy=='TOTAL':
                    tb_amt=float(grp_raw_total.get(bs, grp_hier_total.get(bs,0)))
                else:
                    tb_amt=float(grp_raw_ccy.get((bs,ccy), grp_hier_ccy.get((bs,ccy),0)))

            # a manually built / imported rule that resolved to nothing on
            # BOTH sides would otherwise report a trivial 0 vs 0 "MATCH" —
            # actively misleading in an audit tool, so call it out instead.
            if tb_components is not None and not tb_components and not evidence:
                status='UNRESOLVED'; diff=None; pct=None
            else:
                diff=tb_amt-sub_amt; pct=abs(diff)/max(abs(tb_amt),1)
                status='MATCH' if abs(diff)<1e-9 else 'MATCH_WITHIN_TOLERANCE' if abs(diff)<=self.tolerance_abs or pct<=self.tolerance_pct else 'REVIEW_REQUIRED'
            out.append({'bs_mapping':bs,'currency':ccy,'tb_amount':tb_amt,'submission_amount':sub_amt,'difference':diff,
                        'variance_pct':pct,'status':status,'rule_type':rule.get('rule_type','DIRECT'),
                        'source':rule.get('source','MANUAL'),'evidence':json.dumps(evidence,ensure_ascii=False),
                        'tb_evidence':json.dumps(tb_evidence,ensure_ascii=False) if tb_evidence is not None else None})
        return pd.DataFrame(out)

    # ---------- Manual match builder + save/load mapping support ----------

    def resolve_rule_template(self, template, tb, submissions):
        """Re-resolve a portable rule template (stable keys only: account /
        description / bs_mapping on the TB side, line description / sheet /
        currency on the submission side) against a *current* TB + submissions
        session. Used both right after building a manual match (trivial,
        exact resolution) and when a saved mapping file is re-uploaded against
        a new period's data (fuzzy, best-effort — every miss/low-confidence
        hit is reported as a warning instead of silently dropped)."""
        warnings=[]
        tb_components=[]
        for item in template.get('tb_side',[]) or []:
            acct=str(item.get('account') or '').strip()
            desc=str(item.get('account_desc') or '').strip()
            bsmap=item.get('bs_mapping'); ccy=item.get('currency')
            cand=tb
            if acct:
                hit=tb[tb.account.astype(str).str.strip()==acct]
                if len(hit): cand=hit
                else: cand=tb.iloc[0:0]
            if not len(cand) and desc:
                pool=tb[tb.bs_mapping==bsmap] if bsmap and (tb.bs_mapping==bsmap).any() else tb
                best_idx,best_score=None,0.0
                for idx,row in pool.iterrows():
                    sc=text_similarity(desc,row.account_desc)
                    if sc>best_score: best_idx,best_score=idx,sc
                cand = tb.loc[[best_idx]] if best_idx is not None and best_score>=0.55 else tb.iloc[0:0]
            if ccy and len(cand):
                ccy_hit=cand[cand.tran_ccy==ccy]
                if len(ccy_hit): cand=ccy_hit
            if not len(cand):
                warnings.append(f"TB item not found: account={acct or '—'!s} description={desc or '—'!s}")
                continue
            for _,row in cand.iterrows():
                tb_components.append({'account':row.account,'account_desc':row.account_desc,'bs_mapping':row.bs_mapping,
                                       'currency':row.tran_ccy,'amount':float(row.adjusted_balance),'sign':item.get('sign',1)})

        sub_components=[]
        for item in template.get('submission_side',[]) or []:
            text=item.get('match_text') or item.get('line_description') or ''
            sheet=item.get('sheet'); ccy=item.get('currency'); want_total=item.get('is_total')
            pool=submissions
            if sheet is not None and len(pool):
                p2=pool[pool.sheet==sheet]
                if len(p2): pool=p2
            if ccy and len(pool):
                p2=pool[pool.currency==ccy]
                if len(p2): pool=p2
            if want_total and len(pool):
                p2=pool[pool.is_total==True]
                if len(p2): pool=p2
            if not len(pool):
                warnings.append(f"Submission item not found (no candidate rows): {text!r}")
                continue
            best_idx,best_score=None,0.0
            for idx,row in pool.iterrows():
                sc=text_similarity(text,row.line_description)
                if sc>best_score: best_idx,best_score=idx,sc
            if best_idx is None or best_score<0.5:
                warnings.append(f"Submission item low-confidence match, please review: {text!r} (best match score {best_score:.2f})")
                continue
            row=submissions.loc[best_idx]
            sub_components.append({'submission_file':row.submission_file,'sheet':row.sheet,'row_number':int(row.row_number),
                                    'source_cell':row.source_cell,'line_description':row.line_description,'currency':row.currency,
                                    'amount':float(row.normalized_amount),'multiplier':1,'sign':item.get('sign',1),
                                    'match_confidence':round(best_score,3)})
            if best_score<0.85:
                warnings.append(f"Submission item matched with moderate confidence ({best_score:.2f}): {text!r} → {row.line_description!r}")

        # the new period's submission file might not carry the same reporting
        # scale as when this mapping was first built (or the scale was never
        # in the header text to begin with) — check numerically before
        # reporting a variance that's really just a unit mismatch.
        tb_sum=sum(c['amount']*c.get('sign',1) for c in tb_components)
        sub_raw_sum=sum(c['amount']*c.get('sign',1) for c in sub_components)
        if tb_components and sub_components:
            fit=self.best_scale_fit(sub_raw_sum, tb_sum, self.tolerance_abs, self.tolerance_pct)
            if fit:
                scale_f,sign_f,_=fit
                if scale_f!=1.0 or sign_f!=1.0:
                    for c in sub_components:
                        c['multiplier']=c.get('multiplier',1)*scale_f
                        c['sign']=c.get('sign',1)*sign_f
                    warnings.append(f"Submission values didn't reconcile at face value; applied a "
                                     f"{'×' if scale_f>=1 else '÷'}{scale_f if scale_f>=1 else round(1/scale_f)} scale"
                                     f"{' and a sign flip' if sign_f<0 else ''} automatically — please double-check.")

        resolved={'bs_mapping':template.get('label') or 'Imported match','label':template.get('label'),
                  'currency':template.get('currency','TOTAL'),'rule_type':template.get('rule_type','MANUAL_CLUB'),
                  'source':'IMPORTED','tb_components':tb_components,'components':sub_components}
        return resolved, warnings

    # ---------- Lineage ----------

    def build_lineage(self, tb_tree, recon_rows):
        by_key=defaultdict(list)
        for r in recon_rows:
            by_key[r.get('bs_mapping')].append(r)
        def attach(node, path):
            full = path if node['name']=='(Direct)' else path+[node['name']]
            key=' / '.join(full)
            node['matches']=by_key.get(key) or by_key.get(node['name']) or []
            for ch in node.get('children') or []:
                attach(ch, full)
            return node
        return [attach(g, []) for g in tb_tree]

    def flatten_lineage(self, tree, path=None):
        path=path or []; rows=[]
        for node in tree:
            cur = path if node['name']=='(Direct)' else path+[node['name']]
            matches=node.get('matches') or []
            children=node.get('children') or []
            if matches:
                for m in matches:
                    rows.append({'hierarchy':' / '.join(cur),'level':node['level'],'currency':m.get('currency'),
                                 'tb_amount':m.get('tb_amount'),'submission_amount':m.get('submission_amount'),
                                 'difference':m.get('difference'),'status':m.get('status'),'source':m.get('source'),
                                 'rule_type':m.get('rule_type'),'evidence':m.get('evidence'),'tb_evidence':m.get('tb_evidence')})
            elif not children:
                totals=node.get('currency_totals') or {}
                rows.append({'hierarchy':' / '.join(cur),'level':node['level'],'currency':'TOTAL',
                             'tb_amount':totals.get('TOTAL'),'submission_amount':None,'difference':None,
                             'status':'NO_RULE','source':None,'rule_type':None,'evidence':None})
            if children:
                rows.extend(self.flatten_lineage(children, cur).to_dict('records'))
        return pd.DataFrame(rows)

    # ---------- Export ----------

    def _style_workbook(self, path):
        wb=load_workbook(path)
        navy='173A5E'
        for ws in wb.worksheets:
            ws.freeze_panes='A2'; ws.sheet_view.showGridLines=False
            try: ws.auto_filter.ref=ws.dimensions
            except Exception: pass
            for c in ws[1]: c.fill=PatternFill('solid',fgColor=navy); c.font=Font(color='FFFFFF',bold=True); c.alignment=Alignment(vertical='center')
            for col in ws.columns:
                letter=get_column_letter(col[0].column); width=min(42,max(12,max(len(str(x.value or '')) for x in col)+2)); ws.column_dimensions[letter].width=width
            for row in ws.iter_rows(min_row=2):
                for c in row:
                    if isinstance(c.value,(int,float)): c.number_format='#,##0.00;[Red](#,##0.00);-'
        wb.save(path); return path

    def export_submissions(self, submissions, path=None):
        """A clean, simplified workbook of what was actually extracted from
        the submission sheets — the raw tagged lines plus a currency pivot —
        so the user can hand it to someone else or archive it independently
        of the full audit workbook."""
        path=path or tempfile.mktemp(suffix='.xlsx')
        df=submissions.drop(columns=['hierarchy'],errors='ignore').copy() if submissions is not None else pd.DataFrame()
        cols=['submission_file','sheet','hierarchy_path','line_description','is_total','currency','raw_amount','scale','normalized_amount','source_cell','row_number']
        cols=[c for c in cols if c in df.columns]
        lines=df[cols] if len(df) else pd.DataFrame([{'Info':'No submission lines extracted yet.'}])
        with pd.ExcelWriter(path,engine='openpyxl') as w:
            lines.to_excel(w,sheet_name='Submission Lines',index=False)
            if len(df):
                pivot=pd.pivot_table(df,index=['submission_file','sheet','hierarchy_path','line_description'],
                                      columns='currency',values='normalized_amount',aggfunc='sum',fill_value=0).reset_index()
                pivot.columns=[str(c) for c in pivot.columns]
                pivot.to_excel(w,sheet_name='Summary by Currency',index=False)
        return self._style_workbook(path)

    def export(self, tb, pivot, submissions=None, recon=None, suggestions=None, lineage=None, path=None):
        path=path or tempfile.mktemp(suffix='.xlsx')
        with pd.ExcelWriter(path,engine='openpyxl') as w:
            summary=pd.DataFrame([{'Metric':'TB rows','Value':len(tb)},{'Metric':'TB adjusted balance','Value':tb.adjusted_balance.sum()},
                                   {'Metric':'BS Mapping groups','Value':tb.bs_mapping.nunique()},
                                   {'Metric':'Submission extracted lines','Value':0 if submissions is None else len(submissions)},
                                   {'Metric':'Reconciliation rules','Value':0 if recon is None else len(recon)}])
            summary.to_excel(w,sheet_name='01 Executive Summary',index=False); pivot.to_excel(w,sheet_name='02 TB Pivot',index=False); tb.to_excel(w,sheet_name='03 TB Account Detail',index=False)
            if submissions is not None and len(submissions): submissions.drop(columns=['hierarchy'],errors='ignore').to_excel(w,sheet_name='04 Submission Inventory',index=False)
            if recon is not None and len(recon): recon.to_excel(w,sheet_name='05 Recon Results',index=False)
            if suggestions: pd.DataFrame(suggestions).drop(columns=['components','match_basis'],errors='ignore').to_excel(w,sheet_name='06 Auto Suggestions',index=False)
            if lineage is not None and len(lineage): lineage.drop(columns=['evidence','tb_evidence'],errors='ignore').to_excel(w,sheet_name='07 Reconciliation Map',index=False)
        return self._style_workbook(path)
