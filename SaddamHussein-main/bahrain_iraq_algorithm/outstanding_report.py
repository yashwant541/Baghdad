"""Outstanding Report -> the two PivotTables of the Off-Balance specification, as code.

  Pivot 1  rows: LC/GTEE DESC   columns: BILL_CCY   (no CATEGORY, no filter)    values: Sum of Equ-IQD
  Pivot 2  rows: CATEGORY > LC/GTEE DESC > BILL_CCY      columns: Bucket        values: Sum of Equ-IQD

The raw-data sheet and its header row are found dynamically (nothing is tied to a sheet name, a
header row number, an Excel column letter, or a currency / category / bucket value); the five
mandatory fields are standardised; Equ-IQD is cleaned without ever turning bad values into zero;
both pivots carry grand totals that are reconciled to the cleaned source.

Python 3.9 compatible.
"""
import math
import re

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BLANK = "(blank)"
GRAND = "Grand Total"
MANDATORY = ("CATEGORY", "BILL_CCY", "LC_GTEE_DESC", "EQUI_IQD", "BUCKET")
DISPLAY = {"CATEGORY": "CATEGORY", "BILL_CCY": "BILL_CCY", "LC_GTEE_DESC": "LC/GTEE DESC",
           "EQUI_IQD": "Equ-IQD", "BUCKET": "Bucket"}
MEASURE = "Sum of Equ-IQD"

# every raw-data field the report is known to carry
RAW_FIELDS = ["CATEGORY", "CUST_SEGMENT", "SECTION_ID", "SEGMENT NAME (SEC ID)", "CUSTOMER_ID", "CUSTOMER NAME",
              "CUSTOMER_REF_NO", "DEAL_NUMBER", "STEP_ID", "PROCESS DATE", "BILL_CCY", "BILL_AMOUNT", "DEAL_CCY",
              "DEAL_BAL", "SIGHT(Y/n)", "AVAILISATION (Y/N)", "FINANCED", "FIN_TYPE", "STATUS", "FIN STATUS DESC",
              "ACCEPTED MATURITY DATE", "FIN_START_DATE", "FIN_DUE_DATE", "FINANCE_NUMBER", "FIN_OUS_CCY",
              "FIN_OUSTANDING AMT", "MAR_CURRENCY", "MARGIN BALANCE", "EXP_DUE_DATE/MATURITY_DATE", "Days o/s",
              "CCIS_ID", "DOC_PREP", "CONFIRM_Y_N", "CONFIRM_AMOUNT", "DISCREPANT DOC", "DOC_IN_ORDER", "ISB", "RMT",
              "COL", "ADV", "RMB", "APB", "BENE/DRAWER", "APPLICANT/DRAWEE", "ACTUAL APPLICANT", "LC/GTEE DESC",
              "Equ-IQD", "Bucket", "Rate"]


def _key(v):
    """'Equ-IQD ', 'EQU IQD', 'Equ_IQD' -> 'equiqd'; 'LC / GTEE DESC' -> 'lcgteedesc'"""
    return re.sub(r"[^a-z0-9]", "", str(v if v is not None else "").lower())


RECOGNIZED = set(_key(f) for f in RAW_FIELDS) | {"equiiqd", "equivalentiqd", "billcurrency", "lcgteedescription", "exchangerate"}
ALIASES = {"CATEGORY": {"category"},
           "BILL_CCY": {"billccy", "billcurrency"},
           "LC_GTEE_DESC": {"lcgteedesc", "lcgteedescription"},
           "EQUI_IQD": {"equiqd", "equiiqd", "equivalentiqd"},
           "BUCKET": {"bucket"}}
NULL_TEXT = {"", "nan", "none", "null", "nat", "n/a"}
PIVOT_NAME_HINTS = re.compile(r"pivot|summary|rate|exchange|fx|total", re.I)


class OutstandingError(ValueError):
    """A problem the user can act on (missing field, unknown category, empty data...)."""


# ------------------------------------------------------------------------------ normalisation

def normalize_header(v):
    """-> (display label, comparison key). The display label keeps the original text (trimmed, single
    spaces); the key ignores case, spaces and punctuation."""
    label = re.sub(r"\s+", " ", str(v if v is not None else "")).strip()
    return label, _key(label)


def normalize_group_value(v):
    """The label exactly as it is written in the sheet; only a truly empty cell is None (shown as (blank))."""
    if v is None:
        return None
    if isinstance(v, float):
        if math.isnan(v):
            return None
        if v.is_integer():
            return str(int(v))
    s = str(v)
    if s == "":
        return None
    return "(spaces only)" if s.strip() == "" else s       # a cell holding only spaces is NOT an empty cell: Excel lists it apart from (blank)


def clean_numeric_field(value):
    """-> (number or None, status) with status 'ok' | 'blank' | 'invalid' - the way an Excel PivotTable sees it: a number
    is summed, an empty cell or text is not (text is never converted and never turned into 0)."""
    if value is None:
        return None, "blank"
    if isinstance(value, bool):
        return None, "invalid"
    if isinstance(value, (int, float, np.integer, np.floating)):
        f = float(value)
        if math.isnan(f):
            return None, "blank"                      # an empty cell
        return (f, "ok") if math.isfinite(f) else (None, "invalid")
    return (None, "blank") if str(value).strip() == "" else (None, "invalid")


def loose_number(value):
    """For information only (never used in a total): the number a text such as '1,234.5' or '(250)' looks like, else None."""
    s = str(value).strip().translate(str.maketrans("٠١٢٣٤٥٦٧٨٩٬٫", "0123456789,."))
    t = re.sub(r"[,\s$€£¥]", "", s)
    neg = False
    if t.startswith("(") and t.endswith(")"):
        neg, t = True, t[1:-1]
    if t.endswith("-") and t.count("-") == 1:
        neg, t = True, t[:-1]
    if re.fullmatch(r"[-+]?\d+(\.\d+)?", t):
        return -float(t) if neg else float(t)
    return None


# ------------------------------------------------------------------------------ sheet + header detection

def detect_header_row(raw, max_rows=20):
    """Scan the first rows: the header is the row naming the most recognised fields, accepted only when
    all five mandatory fields resolve. -> dict(row, resolved{STD: column index}, recognised) or None."""
    best = None
    for i in range(min(max_rows, len(raw))):
        keys = {}
        for j, v in enumerate(raw.iloc[i].tolist()):
            if v is None or (isinstance(v, float) and math.isnan(v)) or str(v).strip() == "":
                continue
            keys.setdefault(_key(v), j)
        n_rec = sum(1 for k in keys if k in RECOGNIZED)
        resolved = {}
        for std, aliases in ALIASES.items():
            for a in aliases:
                if a in keys:
                    resolved[std] = keys[a]
                    break
        cand = {"row": i, "resolved": resolved, "recognised": n_rec, "complete": len(resolved) == len(MANDATORY)}
        if best is None or (cand["complete"], n_rec) > (best["complete"], best["recognised"]):
            best = cand
    if best and best["complete"]:
        return best
    return None


def score_source_sheet(name, raw):
    """Higher = more likely to be the raw transaction sheet. Pivot output, summary and exchange-rate
    sheets lack the mandatory header and/or are recognisable by their labels, and score nothing."""
    head = detect_header_row(raw)
    info = {"sheet": name, "score": 0, "header": None, "reason": ""}
    first_cells = " ".join(str(v) for v in raw.iloc[:20].to_numpy().ravel().tolist() if v is not None)[:4000]
    if head is None:
        info["reason"] = "mandatory fields not found in the first 20 rows"
        return info
    if re.search(r"row labels|sum of ", first_cells, re.I) and head["recognised"] < 10:
        info["reason"] = "looks like a PivotTable output"
        return info
    body = raw.iloc[head["row"] + 1:]
    detail_rows = int((body.notna().sum(axis=1) >= 3).sum())
    penalty = 0.5 if PIVOT_NAME_HINTS.search(name) else 1.0
    info.update(header=head, detail_rows=detail_rows, score=(head["recognised"] * 100000 + detail_rows) * penalty)
    return info


def select_source_sheet(candidates, requested=None):
    if requested:
        for c in candidates:
            if c["sheet"] == requested:
                if c["header"] is None:
                    raise OutstandingError("Sheet %r does not contain the mandatory fields (%s)." %
                                           (requested, ", ".join(DISPLAY[m] for m in MANDATORY)))
                return c
        raise OutstandingError("Sheet %r was not found. Sheets: %s" % (requested, ", ".join(c["sheet"] for c in candidates)))
    valid = [c for c in candidates if c["header"] is not None and c["score"] > 0]
    if not valid:
        reasons = "; ".join("%s: %s" % (c["sheet"], c["reason"]) for c in candidates)
        raise OutstandingError("No raw-data sheet with the mandatory fields (%s) was found. %s" %
                               (", ".join(DISPLAY[m] for m in MANDATORY), reasons))
    return max(valid, key=lambda c: c["score"])


# ------------------------------------------------------------------------------ reading + cleaning

def read_source_data(path, sheet=None):
    """-> (dataframe with original column labels, info). Every sheet is examined; the best raw-data
    sheet and its header row are chosen from the data itself."""
    xls = pd.ExcelFile(path, engine="openpyxl")
    candidates, raws = [], {}
    for name in xls.sheet_names:
        # keep_default_na=False: pandas would otherwise turn the text "N/A" / "NA" / "null" into an empty cell
        raw = xls.parse(name, header=None, dtype=object, keep_default_na=False, na_values=[""])
        raws[name] = raw
        candidates.append(score_source_sheet(name, raw))
    chosen = select_source_sheet(candidates, sheet)
    raw = raws[chosen["sheet"]]
    head = chosen["header"]
    labels, used = [], {}
    for j, v in enumerate(raw.iloc[head["row"]].tolist()):
        label, _ = normalize_header(v)
        label = label or "Unnamed %d" % (j + 1)
        if label in used:
            used[label] += 1
            label = "%s.%d" % (label, used[label])
        else:
            used[label] = 0
        labels.append(label)
    df = raw.iloc[head["row"] + 1:].copy()
    df.columns = labels
    df = df.reset_index(drop=True)
    df["_SRC_ROW"] = [head["row"] + 2 + i for i in range(len(df))]            # Excel row number of every data row
    info = {"sheet": chosen["sheet"], "header_row": head["row"] + 1, "resolved": dict(head["resolved"]),
            "labels": labels, "candidates": [(c["sheet"], c["score"], c["reason"]) for c in candidates]}
    return df, info


def standardize_columns(df, info):
    """Drop blank rows/columns and repeated header rows; rename the five mandatory fields to
    CATEGORY / BILL_CCY / LC_GTEE_DESC / EQUI_IQD / BUCKET (original labels remembered)."""
    labels = info["labels"]
    rename, original = {}, {}
    for std, j in info["resolved"].items():
        rename[labels[j]] = std
        original[std] = labels[j]
    blank = lambda v: v is None or (isinstance(v, float) and math.isnan(v)) or str(v).strip() == ""
    df = df.loc[:, [c for c in df.columns if c == "_SRC_ROW" or not df[c].map(blank).all()]]
    is_blank_row = df.apply(lambda r: all(blank(v) for k, v in r.items() if k != "_SRC_ROW"), axis=1)
    info["blank_rows_removed"] = int(is_blank_row.sum())
    df = df.loc[~is_blank_row].reset_index(drop=True)
    # a header row repeated inside the data (copy-pasted blocks)
    keys_row = {c: _key(c) for c in rename if c in df.columns}
    if keys_row:
        rep = df.apply(lambda r: sum(1 for c, k in keys_row.items() if _key(r[c]) == k) >= max(3, len(keys_row) - 1), axis=1)
        info["repeated_headers_removed"] = int(rep.sum())
        df = df.loc[~rep].reset_index(drop=True)
    else:
        info["repeated_headers_removed"] = 0
    df = df.rename(columns=rename)
    info["original_labels"] = original
    return df


def clean_source_data(df):
    """A plain pivot source: the grouping labels as written, Equ-IQD read the way Excel reads it. Nothing is corrected.
    Rows that cannot be summed (empty / text Equ-IQD) are kept with a status so they can be listed, and a row that
    looks like a total line typed under the data is only flagged - it is counted, as an Excel pivot would."""
    df = df.copy()
    for col in ("CATEGORY", "BILL_CCY", "LC_GTEE_DESC", "BUCKET"):
        df[col] = df[col].map(normalize_group_value)
    parsed = df["EQUI_IQD"].map(clean_numeric_field)
    df["_EQUI_RAW"] = df["EQUI_IQD"]
    df["EQUI_IQD"] = parsed.map(lambda x: x[0])
    df["_STATUS"] = parsed.map(lambda x: x[1])
    no_group = df[["CATEGORY", "BILL_CCY", "LC_GTEE_DESC", "BUCKET"]].isna().all(axis=1)
    run, looks = 0.0, []
    for i in df.index:
        if df.at[i, "_STATUS"] != "ok":
            looks.append(False)
            continue
        amt = float(df.at[i, "EQUI_IQD"])
        flag = bool(no_group[i] and run and abs(amt - run) <= max(1.0, 1e-6 * abs(run)))
        looks.append(flag)
        if not flag:
            run += amt
    df["_LOOKS_TOTAL"] = looks
    return df


def validate_source_data(df, info, source_name):
    """Counts and checks for the Validation sheet. -> ordered list of (label, value)."""
    ok = df["_STATUS"] == "ok"
    valid = df[ok]
    nblank = lambda col: int(df[col].isna().sum())
    dup_cols = [c for c in ("CUSTOMER_ID", "CUSTOMER_REF_NO", "DEAL_NUMBER", "STEP_ID", "PROCESS DATE", "BILL_CCY",
                            "BILL_AMOUNT", "EQUI_IQD") if c in df.columns]
    dups = int(df.duplicated(subset=dup_cols, keep=False).sum()) if len(dup_cols) >= 3 else 0
    # distinct labels that only differ by case/spacing would show up as duplicate-looking buckets
    bucket_variants = {}
    for b in df["BUCKET"].dropna().unique():
        bucket_variants.setdefault(_key(b), []).append(b)
    looks_dup = [v for v in bucket_variants.values() if len(v) > 1]
    bsum = valid.assign(_b=valid["BUCKET"].fillna(BLANK)).groupby("_b")["EQUI_IQD"].sum()
    return [
        ("Source workbook", source_name),
        ("Selected source sheet", info["sheet"]),
        ("Detected header row", info["header_row"]),
        ("Source rows read", int(info.get("rows_read", len(df)))),
        ("Cleaned detail rows", int(len(df))),
        ("Repeated header rows removed", int(info.get("repeated_headers_removed", 0))),
        ("Rows with blank Equ-IQD (excluded)", int((df["_STATUS"] == "blank").sum())),
        ("Rows with invalid / non-numeric Equ-IQD (excluded)", int((df["_STATUS"] == "invalid").sum())),
        ("Rows that look like a total line typed under the data (counted)", int(df["_LOOKS_TOTAL"].sum())),
        ("Rows with blank CATEGORY", nblank("CATEGORY")),
        ("Rows with blank LC/GTEE DESC", nblank("LC_GTEE_DESC")),
        ("Rows with blank BILL_CCY", nblank("BILL_CCY")),
        ("Rows with blank Bucket", nblank("BUCKET")),
        ("Potential duplicate transaction rows (not removed)", dups),
        ("Source Equ-IQD total (valid rows)", float(valid["EQUI_IQD"].sum())),
        ("Distinct CATEGORY values", ", ".join(sorted(str(x) for x in valid["CATEGORY"].dropna().unique()))),
        ("Distinct BILL_CCY values", ", ".join(sorted(str(x) for x in valid["BILL_CCY"].dropna().unique()))),
        ("Distinct Bucket values", ", ".join(order_pivot_columns(list(bsum.index), "bucket"))),
        ("Bucket labels that differ only by case/spacing", "; ".join(" / ".join(map(repr, v)) for v in looks_dup) or "none"),
        ("Sum of Equ-IQD by Bucket", "; ".join("%s: %s" % (k, format(bsum[k], ",.2f")) for k in order_pivot_columns(list(bsum.index), "bucket"))),
    ]


# ------------------------------------------------------------------------------ ordering

def _bucket_lower_bound(label):
    m = re.search(r"\d[\d,]*(?:\.\d+)?", str(label))
    if not m:
        return None
    v = float(m.group(0).replace(",", ""))
    # "Above 365" / ">365" / "365+" sit after the range that ends at 365
    if re.search(r"above|over|more|greater|>|\+", str(label), re.I) and not re.search(r"^\s*\d", str(label)):
        v += 0.5
    return v


def order_pivot_columns(values, kind):
    """kind 'bucket': ordered by a lower bound read from the label (labels are never sorted as text),
    unnumbered labels keep their order of first appearance, (blank) last. kind 'currency': A-Z, (blank) last."""
    vals = [v for v in values if v != BLANK]
    if kind == "bucket":
        pos = {v: i for i, v in enumerate(vals)}
        vals.sort(key=lambda v: (_bucket_lower_bound(v) is None, _bucket_lower_bound(v) or 0.0, pos[v]))
    else:
        vals.sort(key=lambda v: str(v).lower())
    return vals + ([BLANK] if BLANK in values else [])


def _sorted_text(values):
    vals = [v for v in values if v != BLANK]
    vals.sort(key=lambda v: str(v).lower())
    return vals + ([BLANK] if BLANK in values else [])


# ------------------------------------------------------------------------------ pivots

def _grouped(df):
    d = df[df["_STATUS"] == "ok"].copy()
    for col in ("CATEGORY", "LC_GTEE_DESC", "BILL_CCY", "BUCKET"):
        d[col] = d[col].fillna(BLANK)
    return d


def build_category_bucket_pivot(df):
    """Pivot 1. Returns {'pandas': the pivot_table with margins, 'flat': rows incl. subtotals}."""
    d = _grouped(df)
    if d.empty:
        raise OutstandingError("There are no valid Equ-IQD rows to pivot.")
    pv = pd.pivot_table(d, index=["CATEGORY", "LC_GTEE_DESC", "BILL_CCY"], columns=["BUCKET"], values="EQUI_IQD",
                        aggfunc="sum", fill_value=0, margins=True, margins_name=GRAND, dropna=False)
    buckets = order_pivot_columns(list(d["BUCKET"].unique()), "bucket")
    leaf = d.groupby(["CATEGORY", "LC_GTEE_DESC", "BILL_CCY", "BUCKET"])["EQUI_IQD"].sum()
    cats = _sorted_text(list(d["CATEGORY"].unique()))
    rows = []

    def vec(sub):
        s = sub.groupby("BUCKET")["EQUI_IQD"].sum()
        return [float(s.get(b, 0.0)) for b in buckets]

    for cat in cats:
        dc = d[d["CATEGORY"] == cat]
        v = vec(dc)
        rows.append({"level": "Category", "labels": [cat, "", ""], "ccy": "TOTAL", "values": v, "total": float(sum(v))})
        for lc in _sorted_text(list(dc["LC_GTEE_DESC"].unique())):
            dl = dc[dc["LC_GTEE_DESC"] == lc]
            v = vec(dl)
            rows.append({"level": "LC/GTEE", "labels": [cat, lc, ""], "ccy": "TOTAL", "values": v, "total": float(sum(v))})
            for ccy in _sorted_text(list(dl["BILL_CCY"].unique())):
                v = vec(dl[dl["BILL_CCY"] == ccy])
                rows.append({"level": "Currency", "labels": [cat, lc, ccy], "ccy": ccy, "values": v, "total": float(sum(v))})
    v = vec(d)
    rows.append({"level": "Grand Total", "labels": [GRAND, "", ""], "ccy": "TOTAL", "values": v, "total": float(sum(v))})
    return {"pandas": pv, "grand_total": float(pv.iloc[-1, -1]),
            "flat": {"title": "Equ-IQD by Category, LC/GTEE Description, Billing Currency and Bucket",
                     "label_cols": [DISPLAY["CATEGORY"], DISPLAY["LC_GTEE_DESC"], DISPLAY["BILL_CCY"]],
                     "columns": buckets + [GRAND], "rows": rows, "category": None}}


def build_lcgtee_currency_pivot(df):
    """Pivot 1: rows LC/GTEE DESC, columns BILL_CCY, values Sum of Equ-IQD. CATEGORY is not used and there is no filter."""
    d = _grouped(df)
    if d.empty:
        raise OutstandingError("There are no numeric Equ-IQD rows to pivot.")
    pv = pd.pivot_table(d, index=["LC_GTEE_DESC"], columns=["BILL_CCY"], values="EQUI_IQD", aggfunc="sum",
                        fill_value=0, margins=True, margins_name=GRAND, dropna=False)
    ccys = order_pivot_columns(list(d["BILL_CCY"].unique()), "currency")
    rows = []
    for lc in _sorted_text(list(d["LC_GTEE_DESC"].unique())):
        s = d[d["LC_GTEE_DESC"] == lc].groupby("BILL_CCY")["EQUI_IQD"].sum()
        v = [float(s.get(c, 0.0)) for c in ccys]
        rows.append({"level": "LC/GTEE", "labels": [lc], "ccy": "TOTAL", "values": v, "total": float(sum(v))})
    s = d.groupby("BILL_CCY")["EQUI_IQD"].sum()
    v = [float(s.get(c, 0.0)) for c in ccys]
    rows.append({"level": "Grand Total", "labels": [GRAND], "ccy": "TOTAL", "values": v, "total": float(sum(v))})
    return {"pandas": pv, "grand_total": float(pv.iloc[-1, -1]),
            "flat": {"title": "Equ-IQD by LC/GTEE Description and Billing Currency",
                     "label_cols": [DISPLAY["LC_GTEE_DESC"]], "columns": ccys + [GRAND], "rows": rows,
                     "category": None, "column_ccy": ccys + ["TOTAL"]}}


# ------------------------------------------------------------------------------ driver

def process_outstanding(path, source_name=None, sheet=None):
    """Read -> standardise -> clean -> validate -> pivot -> reconcile. Raises OutstandingError with a
    clear message rather than returning something silently wrong."""
    source_name = source_name or str(path)
    raw_df, info = read_source_data(path, sheet)
    info["rows_read"] = len(raw_df)
    df = standardize_columns(raw_df, info)
    if df.empty:
        raise OutstandingError("The raw-data sheet %r has no detail rows below its header." % info["sheet"])
    df = clean_source_data(df)
    validation = validate_source_data(df, info, source_name)
    p1 = build_lcgtee_currency_pivot(df)                     # Pivot 1: rows LC/GTEE DESC, columns BILL_CCY
    p2 = build_category_bucket_pivot(df)                     # Pivot 2: rows CATEGORY > LC/GTEE DESC > BILL_CCY, columns Bucket
    src_total = float(df.loc[df["_STATUS"] == "ok", "EQUI_IQD"].sum())
    diff1, diff2 = src_total - p1["grand_total"], src_total - p2["grand_total"]
    warnings = []
    if abs(diff1) > 0.005 or abs(diff2) > 0.005:
        warnings.append("Pivot grand totals do not agree with the cleaned source total (difference %s / %s)." %
                        (format(diff1, ",.2f"), format(diff2, ",.2f")))
    n_tot = int(df["_LOOKS_TOTAL"].sum())
    if n_tot:
        warnings.append("%d row(s) look like a total line typed under the data (no category, description, currency or bucket, and the amount equals "
                        "the sum of the rows above). They ARE counted, as in an Excel PivotTable - see the Row_Exceptions sheet." % n_tot)
    # what Excel itself gives for the raw Equ-IQD column: numbers are summed, empty cells and text are not
    raw_col = info["labels"][info["resolved"]["EQUI_IQD"]]
    is_num = lambda v: isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool) and math.isfinite(float(v))
    excel_total = float(sum(float(v) for v in raw_df[raw_col] if is_num(v)))
    diff_excel = src_total - excel_total
    texts = df[df["_STATUS"] == "invalid"]
    like = [(r["_SRC_ROW"], loose_number(r["_EQUI_RAW"])) for r in texts.to_dict("records")]
    like = [(row, v) for row, v in like if v is not None]
    like_total = float(sum(v for _, v in like))
    n_blank = int((df["_STATUS"] == "blank").sum())
    if len(texts) or n_blank:
        msg = "%d row(s) have an empty Equ-IQD and %d have text in it; they are not in any total (an Excel PivotTable ignores them too)" % (n_blank, len(texts))
        if like:
            msg += ", and %d of the text values read as numbers (%s in all) - see the Row_Exceptions sheet" % (len(like), format(like_total, ",.2f"))
        warnings.append(msg + ".")
    exc = df[(df["_STATUS"] != "ok") | df["_LOOKS_TOTAL"]].copy()
    looked = dict((row, v) for row, v in like)
    exc["_WHY"] = [("COUNTED: looks like a total line typed under the data" if lt else
                    "IGNORED: empty Equ-IQD" if st == "blank" else
                    ("IGNORED: text that reads as the number %s" % format(looked[sr], ",.2f")) if sr in looked else "IGNORED: text, not a number")
                   for st, lt, sr in zip(exc["_STATUS"], exc["_LOOKS_TOTAL"], exc["_SRC_ROW"])]
    validation += [("Rows dropped as completely blank", int(info.get("blank_rows_removed", 0))),
                   ("Excel-style total of the raw Equ-IQD column (numeric cells only)", excel_total),
                   ("Equ-IQD cells ignored because they hold text", int(len(texts))),
                   ("of which read as numbers (not counted)", int(len(like))),
                   ("Amount those would add if they were counted", like_total),
                   ("Difference: pivot total - Excel-style total of the raw column", diff_excel)]
    validation += [("Pivot 1 Grand Total", p1["grand_total"]), ("Pivot 2 Grand Total", p2["grand_total"]),
                   ("Reconciliation difference (source - Pivot 1)", diff1), ("Reconciliation difference (source - Pivot 2)", diff2),
                   ("Warnings", "; ".join(warnings) or "none")]
    return {"info": info, "df": df, "exceptions": exc, "excel_total": excel_total, "pivot1": p1, "pivot2": p2,
            "validation": validation, "warnings": warnings,
            "source_total": src_total}


# ------------------------------------------------------------------------------ depth-search hand-off

def depth_specs(result, maturity_sheet_scope="maturity"):
    """The two pivots as flat sheet specs for the depth search: header row 1, data from row 2, label
    columns first, then one column per bucket / currency, then Grand Total. Each numeric cell becomes a
    search target. Pivot 2 (by Bucket = tenure) is only meaningful against the Maturity sheet."""
    specs = []
    for n, (key, scope) in enumerate((("pivot1", "all"), ("pivot2", maturity_sheet_scope)), 1):
        flat = result[key]["flat"]
        nl = len(flat["label_cols"])
        headers = list(flat["label_cols"]) + list(flat["columns"])
        rows, bold, targets = [], [], []
        for i, r in enumerate(flat["rows"]):
            sheet_row = i + 2
            rows.append(list(r["labels"]) + list(r["values"]) + [r["total"]])
            if r["level"] != "Currency" and not (key == "pivot1" and r["level"] == "LC/GTEE"):
                bold.append(sheet_row)
            vals = list(r["values"]) + [r["total"]]
            for j, amount in enumerate(vals):
                col = nl + 1 + j
                col_name = flat["columns"][j]
                if key == "pivot2":
                    ccy = r["ccy"]
                    label = " / ".join(x for x in r["labels"] if x) + ("" if col_name == GRAND else "  |  " + col_name)
                    names = [x for x in r["labels"] if x]
                else:
                    ccy = flat["column_ccy"][j]
                    label = (" / ".join(x for x in r["labels"] if x)) + ("" if col_name == GRAND else "  |  " + col_name)
                    names = [x for x in r["labels"] if x and x != GRAND]
                targets.append({"row": sheet_row, "col": col, "level": "OB " + r["level"], "label": label,
                                "currency": "TOTAL" if (ccy in (BLANK, "") or ccy is None) else ccy,
                                "amount": float(amount), "names": names})
        specs.append({"sheet": "OB Pivot %d" % n, "title": flat["title"], "headers": headers, "rows": rows,
                      "n_label": nl, "bold": bold, "scope": scope, "targets": targets,
                      "category": flat.get("category")})
    return specs


# ------------------------------------------------------------------------------ output workbook

HEADER_FILL = PatternFill("solid", start_color="1F3864", end_color="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF")
NUM_FMT = '#,##0;[Red](#,##0);"-"'
TOP_BORDER = Border(top=Side(style="medium", color="1F3864"))


def _autosize(ws, max_width=48):
    widths = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None:
                widths[c.column_letter] = max(widths.get(c.column_letter, 0), len(str(c.value)) + 2)
    for col, w in widths.items():
        ws.column_dimensions[col].width = min(max(w, 9), max_width)


def write_pivot_sheet(ws, flat, header_row=3, subtitle=None):
    ws["A1"] = flat["title"]
    ws["A1"].font = Font(bold=True, size=13, color="1F3864")
    if subtitle:
        ws["A2"] = subtitle
        ws["A2"].font = Font(italic=True)
    heads = list(flat["label_cols"]) + list(flat["columns"])
    for j, h in enumerate(heads, 1):
        c = ws.cell(row=header_row, column=j, value=h)
        c.fill, c.font = HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(horizontal="center" if j > len(flat["label_cols"]) else "left", vertical="center", wrap_text=True)
    nl = len(flat["label_cols"])
    for i, r in enumerate(flat["rows"], header_row + 1):
        bold = r["level"] in ("Category", "Grand Total") or (r["level"] == "LC/GTEE" and nl == 3)
        indent = {"Category": 0, "LC/GTEE": 1, "Currency": 2}.get(r["level"], 0) if nl == 3 else 0
        for j, lab in enumerate(r["labels"], 1):          # labels repeat down the hierarchy so filters work
            c = ws.cell(row=i, column=j, value=lab)
            c.alignment = Alignment(indent=indent if j == 1 else 0)
            if bold:
                c.font = Font(bold=True)
        for j, v in enumerate(list(r["values"]) + [r["total"]], nl + 1):
            c = ws.cell(row=i, column=j, value=v)
            c.number_format = NUM_FMT
            if bold or j == nl + len(flat["columns"]):
                c.font = Font(bold=True)
        if r["level"] == "Grand Total":
            for j in range(1, nl + len(flat["columns"]) + 1):
                ws.cell(row=i, column=j).border = TOP_BORDER
    ws.freeze_panes = ws.cell(row=header_row + 1, column=nl + 1)
    ws.auto_filter.ref = "A%d:%s%d" % (header_row, get_column_letter(nl + len(flat["columns"])), header_row + len(flat["rows"]))
    _autosize(ws)


def write_validation_sheet(ws, validation):
    ws["A1"] = "Validation"
    ws["A1"].font = Font(bold=True, size=13, color="1F3864")
    for j, h in enumerate(("Check", "Value"), 1):
        c = ws.cell(row=3, column=j, value=h)
        c.fill, c.font = HEADER_FILL, HEADER_FONT
    for i, (k, v) in enumerate(validation, 4):
        ws.cell(row=i, column=1, value=k).font = Font(bold=True)
        c = ws.cell(row=i, column=2, value=v)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if isinstance(v, float):
            c.number_format = "#,##0.00;[Red]-#,##0.00"
    ws.column_dimensions["A"].width = 56
    ws.column_dimensions["B"].width = 110


def create_output_workbook(result, out_path):
    """Source_Cleaned, Pivot_LCGTEE_Currency (Pivot 1), Pivot_Category_Bucket (Pivot 2), Validation."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Source_Cleaned"
    df = result["df"]
    orig = result["info"].get("original_labels", {})
    cols = [c for c in df.columns if c not in ("_EQUI_RAW", "_STATUS")]
    ws.append([("Source row" if c == "_SRC_ROW" else orig.get(c, c)) for c in cols] + ["Row status"])
    for c in ws[1]:
        c.fill, c.font = HEADER_FILL, HEADER_FONT
    status_text = {"ok": "", "blank": "IGNORED: empty Equ-IQD", "invalid": "IGNORED: text in Equ-IQD"}
    for rec, st, raw_amt in zip(df[cols].itertuples(index=False), df["_STATUS"], df["_EQUI_RAW"]):
        row = [None if (isinstance(v, float) and math.isnan(v)) else v for v in rec]
        if st != "ok":
            row[cols.index("EQUI_IQD")] = raw_amt
        ws.append(row + [status_text[st]])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    _autosize(ws, 32)

    write_pivot_sheet(wb.create_sheet("Pivot_LCGTEE_Currency"), result["pivot1"]["flat"])
    write_pivot_sheet(wb.create_sheet("Pivot_Category_Bucket"), result["pivot2"]["flat"])
    ex = result.get("exceptions")
    if ex is not None and len(ex):
        we = wb.create_sheet("Row_Exceptions")
        we.append(["Source row", "Why", "Equ-IQD as found", "Equ-IQD used", "CATEGORY", "LC/GTEE DESC", "BILL_CCY", "Bucket"])
        for c in we[1]:
            c.fill, c.font = HEADER_FILL, HEADER_FONT
        for d in ex.to_dict("records"):
            we.append([d.get("_SRC_ROW"), d["_WHY"], None if d["_EQUI_RAW"] is None else str(d["_EQUI_RAW"]),
                       None if (isinstance(d["EQUI_IQD"], float) and math.isnan(d["EQUI_IQD"])) else d["EQUI_IQD"],
                       d["CATEGORY"], d["LC_GTEE_DESC"], d["BILL_CCY"], d["BUCKET"]])
        _autosize(we, 44)
        we.freeze_panes = "A2"
    write_validation_sheet(wb.create_sheet("Validation"), result["validation"])
    wb.save(out_path)
    return out_path
