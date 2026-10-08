"""Outstanding Report -> the two PivotTables of the Off-Balance specification, as code.

  Pivot 1  rows: LC/GTEE DESC   columns: BILL_CCY   (no CATEGORY, no filter)    values: Sum of Equ-IQD
  Pivot 2  rows: CATEGORY > LC/GTEE DESC > BILL_CCY      columns: Bucket        values: Sum of Equ-IQD

The raw-data sheet and its header row are found dynamically (nothing is tied to a sheet name, a
header row number, an Excel column letter, or a currency / category / bucket value); the five
mandatory fields are standardised; Equ-IQD is cleaned without ever turning bad values into zero;
both pivots carry grand totals that are reconciled to the cleaned source.

Python 3.9 compatible.
"""
import datetime
import json
import math
import os
import re

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BLANK = "(blank)"
GRAND = "Grand Total"
FRX_COL = "Forex (non-IQD)"       # Pivot 1: the sum of every currency except IQD, next to the Grand Total
LOCAL = "IQD"
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

    def with_forex(v):
        """the currency columns, then the Forex column (every currency except IQD); the total counts each amount once"""
        return v + [float(sum(x for c, x in zip(ccys, v) if c != LOCAL))]

    for lc in _sorted_text(list(d["LC_GTEE_DESC"].unique())):
        s = d[d["LC_GTEE_DESC"] == lc].groupby("BILL_CCY")["EQUI_IQD"].sum()
        v = [float(s.get(c, 0.0)) for c in ccys]
        rows.append({"level": "LC/GTEE", "labels": [lc], "ccy": "TOTAL", "values": with_forex(v), "total": float(sum(v))})
    s = d.groupby("BILL_CCY")["EQUI_IQD"].sum()
    v = [float(s.get(c, 0.0)) for c in ccys]
    rows.append({"level": "Grand Total", "labels": [GRAND], "ccy": "TOTAL", "values": with_forex(v), "total": float(sum(v))})
    return {"pandas": pv, "grand_total": float(pv.iloc[-1, -1]),
            "flat": {"title": "Equ-IQD by LC/GTEE Description and Billing Currency",
                     "label_cols": [DISPLAY["LC_GTEE_DESC"]], "columns": ccys + [FRX_COL, GRAND], "rows": rows,
                     "category": None, "column_ccy": ccys + ["FRX", "TOTAL"]}}


# ------------------------------------------------------------------------------ driver

ROLE_STD = {"category": "CATEGORY", "item": "LC_GTEE_DESC", "currency": "BILL_CCY", "bucket": "BUCKET", "amount": "EQUI_IQD"}
DEFINITION_FILE = "pivot_definition.json"


def definition_path():
    """Where the saved pivot definition lives: next to this module in the library (override: BAHRAIN_IRAQ_PIVOT_DEFINITION)."""
    return os.environ.get("BAHRAIN_IRAQ_PIVOT_DEFINITION") or os.path.join(os.path.dirname(os.path.abspath(__file__)), DEFINITION_FILE)


def load_definition():
    """The saved definition (dict) or None."""
    try:
        with open(definition_path(), encoding="utf-8") as fh:
            d = json.load(fh)
        return d if isinstance(d, dict) and d.get("pivot1") and d.get("pivot2") else None
    except (OSError, ValueError):
        return None


def save_definition(spec):
    """Write the definition into the library. Raises OSError when the folder is read-only."""
    doc = {"format": "bahrain-iraq-outstanding-pivots", "version": 1, "saved": datetime.datetime.now().isoformat(timespec="seconds"),
           "sheet": spec.get("sheet"), "header_row": spec.get("header_row"), "pivot1": spec["pivot1"], "pivot2": spec["pivot2"]}
    path = definition_path()
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(doc, fh, ensure_ascii=False, indent=1)
    return path


def delete_definition():
    try:
        os.remove(definition_path())
        return True
    except OSError:
        return False


def _blank_cell(v):
    return v is None or (isinstance(v, float) and math.isnan(v)) or str(v).strip() == ""


def _is_num(v):
    return isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool) and math.isfinite(float(v))


def _unique_labels(row_values):
    labels, used = [], {}
    for j, v in enumerate(row_values):
        label, _ = normalize_header(v)
        label = label or "Unnamed %d" % (j + 1)
        if label in used:
            used[label] += 1
            label = "%s.%d" % (label, used[label])
        else:
            used[label] = 0
        labels.append(label)
    return labels


def _guess_header(raw):
    head = detect_header_row(raw)
    if head:
        return head["row"]
    best, bn = 0, -1
    for i in range(min(20, len(raw))):
        n = sum(1 for v in raw.iloc[i].tolist() if not _blank_cell(v))
        if n > bn:
            best, bn = i, n
    return best


def _read_raw(path, sheet):
    # keep_default_na=False: pandas would otherwise turn the text "N/A" / "NA" / "null" into an empty cell
    return pd.ExcelFile(path, engine="openpyxl").parse(sheet, header=None, dtype=object, keep_default_na=False, na_values=[""])


def read_sheet_table(path, sheet, header_row=None):
    """Any sheet as a table: -> (DataFrame with the original column labels and a _SRC_ROW column, info)."""
    xls = pd.ExcelFile(path, engine="openpyxl")
    if sheet not in xls.sheet_names:
        raise OutstandingError("Sheet %r was not found. Sheets: %s" % (sheet, ", ".join(xls.sheet_names)))
    raw = _read_raw(path, sheet)
    if not len(raw):
        raise OutstandingError("Sheet %r is empty." % sheet)
    hr = (int(header_row) - 1) if header_row else _guess_header(raw)
    if hr < 0 or hr >= len(raw):
        raise OutstandingError("Header row %s is outside sheet %r (%d rows)." % (header_row, sheet, len(raw)))
    labels = _unique_labels(raw.iloc[hr].tolist())
    df = raw.iloc[hr + 1:].copy()
    df.columns = labels
    df = df.reset_index(drop=True)
    df["_SRC_ROW"] = [hr + 2 + i for i in range(len(df))]
    return df, {"sheet": sheet, "header_row": hr + 1, "labels": labels}


def list_sheets(path):
    """Every sheet with the header row it would be read with and its columns (for the pivot builder)."""
    xls = pd.ExcelFile(path, engine="openpyxl")
    out = []
    for name in xls.sheet_names:
        raw = _read_raw(path, name)
        if not len(raw):
            out.append({"sheet": name, "header_row": 1, "columns": [], "rows": 0})
            continue
        hr = _guess_header(raw)
        cols = [c for c in _unique_labels(raw.iloc[hr].tolist()) if not c.startswith("Unnamed ")]
        out.append({"sheet": name, "header_row": hr + 1, "columns": cols, "rows": int(max(len(raw) - hr - 1, 0))})
    return out


def auto_spec(info):
    """The automatic pivot definition, in terms of the columns that were detected."""
    lab = lambda std: info["labels"][info["resolved"][std]]
    return {"version": 1, "sheet": info["sheet"], "header_row": info["header_row"],
            "pivot1": {"item": lab("LC_GTEE_DESC"), "currency": lab("BILL_CCY"), "amount": lab("EQUI_IQD")},
            "pivot2": {"category": lab("CATEGORY"), "item": lab("LC_GTEE_DESC"), "currency": lab("BILL_CCY"),
                       "bucket": lab("BUCKET"), "amount": lab("EQUI_IQD")}}


def _find_label(labels, wanted, what):
    if wanted is None or str(wanted).strip() == "":
        raise OutstandingError("Choose the %s column." % what)
    if wanted in labels:
        return wanted
    k = _key(wanted)
    hits = [l for l in labels if _key(l) == k]
    if hits:
        return hits[0]
    raise OutstandingError("The column %r (%s) is not in this sheet. Columns: %s" % (wanted, what, ", ".join(l for l in labels if not l.startswith("Unnamed "))))


def _role_frame(raw, roles, constants=None):
    """The pivot's source rows under the standard names. roles: STD name -> raw column label (or None)."""
    d = pd.DataFrame({"_SRC_ROW": raw["_SRC_ROW"].values})
    used = [lab for lab in roles.values() if lab]
    keep = ~raw[used].apply(lambda r: all(_blank_cell(v) for v in r), axis=1).values
    for std in ("CATEGORY", "BILL_CCY", "LC_GTEE_DESC", "BUCKET"):
        lab = roles.get(std)
        d[std] = raw[lab].values if lab else (constants or {}).get(std)
    d["EQUI_IQD"] = raw[roles["EQUI_IQD"]].values
    d = d[keep].reset_index(drop=True)
    return clean_source_data(d)


def _formulas_without_value(path, sheet, header_row, amount_label):
    """How many cells of the amount column hold a formula that was never calculated (no saved value) - they read as empty."""
    try:
        wf = load_workbook(path, data_only=False, read_only=True)[sheet]
        wv = load_workbook(path, data_only=True, read_only=True)[sheet]
        hdr = next(wf.iter_rows(min_row=header_row, max_row=header_row, values_only=True))
        labels = _unique_labels(list(hdr))
        if amount_label not in labels:
            return None
        ix = labels.index(amount_label)
        n = 0
        for rf, rv in zip(wf.iter_rows(min_row=header_row + 1, values_only=True), wv.iter_rows(min_row=header_row + 1, values_only=True)):
            f = rf[ix] if ix < len(rf) else None
            v = rv[ix] if ix < len(rv) else None
            if isinstance(f, str) and f.startswith("=") and v is None:
                n += 1
        return n
    except Exception:
        return None


def _pivot_check(n, title, raw, roles, frame, pivot, path, info):
    """What to look at before approving a pivot: its total against the raw column's Excel-style total, and every row left out."""
    amount_label = roles["EQUI_IQD"]
    raw_total = float(sum(float(v) for v in raw[amount_label] if _is_num(v)))
    ok = frame[frame["_STATUS"] == "ok"]
    texts = frame[frame["_STATUS"] == "invalid"]
    blanks = frame[frame["_STATUS"] == "blank"]
    like = [(r["_SRC_ROW"], loose_number(r["_EQUI_RAW"])) for r in texts.to_dict("records")]
    like = [(row, v) for row, v in like if v is not None]
    looked = dict(like)
    ignored = []
    for r in frame.to_dict("records"):
        if r["_STATUS"] == "ok" and not r["_LOOKS_TOTAL"]:
            continue
        why = ("COUNTED: looks like a total line typed under the data" if r["_STATUS"] == "ok" else
               "IGNORED: empty amount" if r["_STATUS"] == "blank" else
               ("IGNORED: text that reads as the number %s" % format(looked[r["_SRC_ROW"]], ",.2f")) if r["_SRC_ROW"] in looked else
               "IGNORED: text, not a number")
        ignored.append({"row": int(r["_SRC_ROW"]), "value": None if r["_EQUI_RAW"] is None else str(r["_EQUI_RAW"]), "why": why})
    label_cols = [(std, roles[std]) for std in ("CATEGORY", "LC_GTEE_DESC", "BILL_CCY", "BUCKET") if roles.get(std)]
    return {"pivot": n, "title": title, "amount_column": amount_label,
            "fields": dict((std, lab) for std, lab in roles.items() if lab),
            "rows_read": int(len(frame)), "rows_summed": int(len(ok)), "ignored_empty": int(len(blanks)), "ignored_text": int(len(texts)),
            "text_like_count": int(len(like)), "text_like_total": float(sum(v for _, v in like)),
            "blank_labels": dict((lab, int(frame[std].isna().sum())) for std, lab in label_cols),
            "raw_total": raw_total, "pivot_total": float(pivot["grand_total"]), "difference": float(pivot["grand_total"]) - raw_total,
            "formulas_without_value": _formulas_without_value(path, info["sheet"], info["header_row"], amount_label) if path else None,
            "ignored_rows": ignored[:60], "ignored_total": len(ignored)}


def _exceptions(frame):
    exc = frame[(frame["_STATUS"] != "ok") | frame["_LOOKS_TOTAL"]].copy()
    like = dict((r["_SRC_ROW"], loose_number(r["_EQUI_RAW"])) for r in exc[exc["_STATUS"] == "invalid"].to_dict("records"))
    exc["_WHY"] = [("COUNTED: looks like a total line typed under the data" if lt else
                    "IGNORED: empty Equ-IQD" if st == "blank" else
                    ("IGNORED: text that reads as the number %s" % format(like[sr], ",.2f")) if like.get(sr) is not None else
                    "IGNORED: text, not a number") for st, lt, sr in zip(exc["_STATUS"], exc["_LOOKS_TOTAL"], exc["_SRC_ROW"])]
    return exc


def _assemble(path, raw_df, info, source_name, f1, f2, roles1, roles2, spec, status):
    """Pivots + checks from the two prepared source frames (the automatic and the custom route share this)."""
    validation = validate_source_data(f2, info, source_name)
    p1 = build_lcgtee_currency_pivot(f1)                     # Pivot 1: rows LC/GTEE DESC, columns BILL_CCY
    p2 = build_category_bucket_pivot(f2)                     # Pivot 2: rows CATEGORY > LC/GTEE DESC > BILL_CCY, columns Bucket
    src_total = float(f2.loc[f2["_STATUS"] == "ok", "EQUI_IQD"].sum())
    checks = [_pivot_check(1, "Pivot 1 - LC/GTEE by currency", raw_df, roles1, f1, p1, path, info),
              _pivot_check(2, "Pivot 2 - category, LC/GTEE and currency by bucket", raw_df, roles2, f2, p2, path, info)]
    warnings = []
    for c in checks:
        t = "Pivot %d" % c["pivot"]
        if c["ignored_empty"] or c["ignored_text"]:
            msg = "%s: %d row(s) have an empty amount and %d have text in it; they are not in the total (an Excel PivotTable ignores them too)" % (
                t, c["ignored_empty"], c["ignored_text"])
            if c["text_like_count"]:
                msg += ", and %d of the text values read as numbers (%s in all)" % (c["text_like_count"], format(c["text_like_total"], ",.2f"))
            warnings.append(msg + ".")
        looks = sum(1 for r in c["ignored_rows"] if r["why"].startswith("COUNTED"))
        if looks:
            warnings.append("%s: %d row(s) look like a total line typed under the data (no labels, amount = the sum above). They ARE counted, as in an "
                            "Excel PivotTable." % (t, looks))
        if c["formulas_without_value"]:
            warnings.append("%s: %d formula cell(s) in the amount column have no saved value (the file was never recalculated in Excel), so they are "
                            "read as empty - open and save the file in Excel once, then upload it again." % (t, c["formulas_without_value"]))
    texts2 = f2[f2["_STATUS"] == "invalid"]
    like2 = [loose_number(v) for v in texts2["_EQUI_RAW"]]
    like2 = [v for v in like2 if v is not None]
    validation += [("Rows dropped as completely blank", int(info.get("blank_rows_removed", 0))),
                   ("Excel-style total of the raw Equ-IQD column (numeric cells only)", checks[1]["raw_total"]),
                   ("Equ-IQD cells ignored because they hold text", int(len(texts2))),
                   ("of which read as numbers (not counted)", int(len(like2))),
                   ("Amount those would add if they were counted", float(sum(like2))),
                   ("Difference: pivot total - Excel-style total of the raw column", checks[1]["difference"]),
                   ("Pivot 1 Grand Total", p1["grand_total"]), ("Pivot 2 Grand Total", p2["grand_total"]),
                   ("Reconciliation difference (source - Pivot 1)", float(f1.loc[f1["_STATUS"] == "ok", "EQUI_IQD"].sum()) - p1["grand_total"]),
                   ("Reconciliation difference (source - Pivot 2)", src_total - p2["grand_total"]),
                   ("Warnings", "; ".join(warnings) or "none")]
    return {"info": info, "df": f2, "data1": f1[f1["_STATUS"] == "ok"].copy(), "data2": f2[f2["_STATUS"] == "ok"].copy(),
            "exceptions": _exceptions(f2), "excel_total": checks[1]["raw_total"], "pivot1": p1, "pivot2": p2,
            "validation": validation, "warnings": warnings, "source_total": src_total, "checks": checks, "spec": spec, "status": status}


def process_outstanding(path, source_name=None, sheet=None):
    """Automatic pivots: read -> detect the sheet, header and the five fields -> pivot. Raises OutstandingError with a
    clear message rather than returning something silently wrong."""
    source_name = source_name or str(path)
    raw_df, info = read_source_data(path, sheet)
    info["rows_read"] = len(raw_df)
    raw_all = raw_df.copy()
    df = standardize_columns(raw_df, info)
    if df.empty:
        raise OutstandingError("The raw-data sheet %r has no detail rows below its header." % info["sheet"])
    df = clean_source_data(df)
    spec = auto_spec(info)
    auto_roles = {"CATEGORY": spec["pivot2"]["category"], "LC_GTEE_DESC": spec["pivot2"]["item"], "BILL_CCY": spec["pivot2"]["currency"],
                  "BUCKET": spec["pivot2"]["bucket"], "EQUI_IQD": spec["pivot2"]["amount"]}
    return _assemble(path, raw_all, info, source_name, df, df, auto_roles,
                     {"LC_GTEE_DESC": spec["pivot1"]["item"], "BILL_CCY": spec["pivot1"]["currency"], "EQUI_IQD": spec["pivot1"]["amount"]},
                     spec, "auto")


def build_from_spec(path, spec, source_name=None, status="draft"):
    """Pivots from a definition: sheet, header row and the column used for each role. spec =
    {sheet, header_row, pivot1: {item, currency, amount}, pivot2: {category (optional), item, currency, bucket, amount}}."""
    source_name = source_name or str(path)
    sheet = spec.get("sheet")
    if not sheet:
        sheet = read_source_data(path)[1]["sheet"]
    raw, info = read_sheet_table(path, sheet, spec.get("header_row"))
    info["rows_read"] = len(raw)
    labels = info["labels"]
    p1s, p2s = spec.get("pivot1") or {}, spec.get("pivot2") or {}
    roles1 = {"LC_GTEE_DESC": _find_label(labels, p1s.get("item"), "Pivot 1 item (LC/GTEE DESC)"),
              "BILL_CCY": _find_label(labels, p1s.get("currency"), "Pivot 1 currency (BILL_CCY)"),
              "EQUI_IQD": _find_label(labels, p1s.get("amount"), "Pivot 1 amount (Equ-IQD)")}
    roles2 = {"CATEGORY": _find_label(labels, p2s.get("category"), "Pivot 2 category") if str(p2s.get("category") or "").strip() else None,
              "LC_GTEE_DESC": _find_label(labels, p2s.get("item"), "Pivot 2 item (LC/GTEE DESC)"),
              "BILL_CCY": _find_label(labels, p2s.get("currency"), "Pivot 2 currency (BILL_CCY)"),
              "BUCKET": _find_label(labels, p2s.get("bucket"), "Pivot 2 bucket"),
              "EQUI_IQD": _find_label(labels, p2s.get("amount"), "Pivot 2 amount (Equ-IQD)")}
    f1 = _role_frame(raw, roles1)
    f2 = _role_frame(raw, roles2, {"CATEGORY": "(all)"})
    if not (f1["_STATUS"] == "ok").any() or not (f2["_STATUS"] == "ok").any():
        raise OutstandingError("There are no numeric amounts in the chosen amount column(s) - check the sheet, the header row and the amount column.")
    clean_spec = {"version": 1, "sheet": sheet, "header_row": info["header_row"],
                  "pivot1": {"item": roles1["LC_GTEE_DESC"], "currency": roles1["BILL_CCY"], "amount": roles1["EQUI_IQD"]},
                  "pivot2": {"category": roles2["CATEGORY"] or "", "item": roles2["LC_GTEE_DESC"], "currency": roles2["BILL_CCY"],
                             "bucket": roles2["BUCKET"], "amount": roles2["EQUI_IQD"]}}
    return _assemble(path, raw, info, source_name, f1, f2, roles1, roles2, clean_spec, status)


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
