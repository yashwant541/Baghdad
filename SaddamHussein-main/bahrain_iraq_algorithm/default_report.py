"""Default-mapping run and its outputs.

* expand_templates   - a template of type FC_COLUMNS ("by currency") becomes one rule per currency column that the
                       target sheet really has, plus one rule that clubs every other foreign currency against the
                       'Other Foreign Currencies' column.
* run_default_mapping- applies every rule to the Trial Balance + extracted submissions, with TB amount, submission
                       amount, variance and status for each.
* write_default_workbook - per submission file: the original sheets with matched cells highlighted, the TB pivot, and a
                       report that states each rule and its variance (same idea as the depth-search output).
* build_pdf          - everything Mapping Studio shows, as a downloadable PDF.

Python 3.9 compatible.
"""
import datetime
import io
import re
import zipfile
from collections import OrderedDict, defaultdict

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

from .depth_search import PALETTE, _col_title, _fill, _sheet_ref, _unique_name, build_layout
from .engine import FRX_CODE, LOCAL_CCY
from .pdf_simple import SimplePDF, clean

STATUS_TEXT = {"MATCH": "Match", "MATCH_WITHIN_TOLERANCE": "Match (within tolerance)",
               "REVIEW_REQUIRED": "Variance - review", "NOT_FOUND": "Not found"}
STATUS_HEX = {"MATCH": "C6EFCE", "MATCH_WITHIN_TOLERANCE": "E2EFDA", "REVIEW_REQUIRED": "F8CBAD", "NOT_FOUND": "FFEB9C"}
STATUS_RGB = {"MATCH": (0.78, 0.94, 0.81), "MATCH_WITHIN_TOLERANCE": (0.89, 0.94, 0.85),
              "REVIEW_REQUIRED": (0.97, 0.80, 0.68), "NOT_FOUND": (1.0, 0.92, 0.61)}


# ------------------------------------------------------------------------------------------ rule wording

def _friendly(patterns):
    out = []
    for p in patterns or []:
        pl = p.lower()
        if "depr" in pl or "dpr" in pl:
            w = "depreciation (Depr / Dpr)"
        elif "software" in pl or "s/w" in pl:
            w = "Capitalised Software"
        else:
            w = re.sub(r"\\s[*+]?", " ", p).replace("\\", "")
        if w not in out:
            out.append(w)
    return out


def _ccy_text(c, excl=None):
    if c == "IQD" or c == LOCAL_CCY:
        return "IQD (local currency)"
    if c == FRX_CODE:
        t = "all non-IQD currencies (forex)"
        if excl:
            t += ", except " + ", ".join(excl) + " (those have their own column)"
        return t
    if c == "TOTAL":
        return "all currencies"
    return c


def describe_group_item(it):
    bsm = str(it.get("bs_mapping") or "")
    name = ("every BS group starting with '%s' (all Assets)" % bsm[:-1]) if bsm.endswith("*") else "BS group %s" % bsm
    bits = [name, _ccy_text(it.get("currency"), it.get("exclude_currencies"))]
    if it.get("sign_filter") == "positive":
        bits.append("positive account balances only")
    elif it.get("sign_filter") == "negative":
        bits.append("negative account balances only")
    if it.get("include_desc"):
        bits.append("only accounts described as " + ", ".join(_friendly(it["include_desc"])))
    if it.get("exclude_desc"):
        bits.append("excluding " + ", ".join(_friendly(it["exclude_desc"])))
    return " | ".join(bits)


def _sub_col_text(ccy):
    if ccy == "IQD":
        return "the IQD column of that row that holds the amount (Residents or Non-Residents)"
    if ccy == FRX_CODE:
        return "the foreign-currency column of that row that holds the amount (Residents or Non-Residents)"
    if ccy == "TOTAL":
        return "Total column of that row"
    return "%s column of that row" % ccy


def describe_rule(t):
    if t.get("rule_type") == "FC_COLUMNS":
        sub = (t.get("submission_side") or [{}])[0]
        return ("BS group%s %s: each currency that has its own column on '%s' is matched to that column; every other "
                "foreign currency is clubbed together and matched to 'Other Foreign Currencies'"
                % ("s" if len(t.get("tb_groups", [])) > 1 else "", " + ".join(t.get("tb_groups", [])), sub.get("sheet")))
    groups = [describe_group_item(x) for x in t.get("tb_side", []) if x.get("level") == "group"]
    accts = [x for x in t.get("tb_side", []) if x.get("level") != "group"]
    parts = list(groups)
    if accts:
        uniq = OrderedDict()
        for x in accts:
            uniq[(x.get("bs_mapping"), x.get("account"), x.get("account_desc"), x.get("currency"))] = 1
        shown = "; ".join("%s %s %s" % (k[0], k[1], k[2]) for k in list(uniq)[:3])
        parts.append("fixed accounts (%d): %s%s" % (len(uniq), shown, " ..." if len(uniq) > 3 else ""))
    tb_text = "  +  ".join(parts) if parts else "(no Trial Balance side)"
    sub = (t.get("submission_side") or [{}])[0]
    sg = sub.get("sign", 1)
    return "%s   ->   %s > %s  [%s%s]" % (tb_text, sub.get("sheet"), sub.get("match_text"), _sub_col_text(sub.get("currency")),
                                            ", sign reversed" if sg == -1 else "")


# ------------------------------------------------------------------------------------------ expansion

def _fc_expand(t, sub):
    """One FC_COLUMNS template -> a rule per direct currency column of the sheet + an 'Other Foreign Currencies' rule."""
    out = []
    s0 = (t.get("submission_side") or [{}])[0]
    sheet = s0.get("sheet")
    on_sheet = sub[sub.sheet == sheet] if (sub is not None and len(sub)) else sub
    codes = sorted(set(on_sheet.currency)) if (on_sheet is not None and len(on_sheet)) else []
    direct = [c for c in codes if c not in ("TOTAL", "IQD", FRX_CODE, "UNSPECIFIED") and not str(c).startswith("TOTAL_")]
    has_other = FRX_CODE in codes
    if not direct and not has_other:
        out.append({"label": t["label"], "currency": "TOTAL", "rule_type": "GROUP_RULE", "tb_side": [],
                    "submission_side": copy_sub(s0, "TOTAL"), "_fc_missing": True, "derived": t.get("derived")})
        return out
    for c in direct:
        out.append({"label": "%s - %s" % (t["label"], c), "currency": c, "rule_type": "GROUP_RULE", "_expanded": True,
                    "derived": t.get("derived"), "tb_side": [grp(g, c) for g in t.get("tb_groups", [])],
                    "submission_side": copy_sub(s0, c)})
    if has_other:
        out.append({"label": "%s - Other Foreign Currencies" % t["label"], "currency": FRX_CODE, "rule_type": "GROUP_RULE",
                    "_expanded": True, "derived": t.get("derived"),
                    "tb_side": [grp(g, FRX_CODE, exclude_currencies=list(direct)) for g in t.get("tb_groups", [])],
                    "submission_side": copy_sub(s0, FRX_CODE)})
    return out


_KINDS = {"ASSET_GROUP": "Assets", "LIABILITY_GROUP": "Liabilities"}


def _sheet_names(kind, sub, override=None):
    """Names of the main / foreign-currency / maturity sheet of a statement ('Assets' or 'Liabilities'), read from the
    extracted sheets when they can be told apart, else the CBI default names."""
    names = {"main": kind, "foreign": kind + " by Foreign Currency", "maturity": kind + " by Maturity"}
    have = sorted(set(sub.sheet)) if (sub is not None and len(sub)) else []
    k = kind.lower()
    for sh in have:
        low = str(sh).lower().strip()
        if not low.startswith(k):
            continue
        if "matur" in low or "period" in low or "tenor" in low:      # 'Liabilities According to Period'
            names["maturity"] = sh
        elif "foreign" in low or "curren" in low:                    # 'Liabilities According to Curren(cy)' - Excel cuts names at 31 characters
            names["foreign"] = sh
        elif low == k:
            names["main"] = sh
    names.update(override or {})
    return names, set(have)


def _asset_group(t, sub=None, notes=None):
    """ASSET_GROUP / LIABILITY_GROUP template (the logic of 'Iraq Changes.docx'; the same for both statements): for the
    row of one BS-mapping group (or groups)
    - main sheet, IQD column 3 or 4 = TB group, IQD only (the subtotal is in ONE of them, not their sum)
    - main sheet, foreign column 5 or 6 = TB group, every non-IQD currency
    - main sheet, total column 2   = TB group, all currencies
    - '<kind> by Foreign Currency'  = TB group currency by currency; the rest against 'Other Foreign Currencies'
    - '<kind> by Maturity', total column = TB group, all currencies
    Fields: label, tb_groups, row (text of the row; row_foreign / row_maturity override it per sheet), include (subset of
    IQD, FRX, TOTAL, FC, MATURITY), sign (-1 for liabilities, whose TB balances are negative), sheets (override names).
    A foreign-currency / maturity sheet that was not extracted is skipped and mentioned in notes."""
    kind = _KINDS.get(t.get("rule_type"), "Assets")
    groups = t.get("tb_groups") or []
    names, have = _sheet_names(kind, sub, t.get("sheets"))
    inc = set(t.get("include") or ["IQD", "FRX", "TOTAL", "FC", "MATURITY"])
    sign = t.get("sign", -1 if kind == "Liabilities" else 1)
    text = t.get("row") or t.get("label")

    def side(sheet, txt, ccy):
        return [{"match_text": txt, "sheet": sheet, "currency": ccy, "is_total": False, "sign": sign}]

    out = []
    for ccy, what in (("IQD", "IQD (%s column 3 or 4)" % kind), ("FRX", "foreign currencies (%s column 5 or 6)" % kind),
                      ("TOTAL", "total (%s column 2)" % kind)):
        if ccy in inc:
            out.append({"label": "%s - %s" % (t["label"], what), "currency": ccy, "rule_type": "GROUP_RULE", "_expanded": True, "derived": True,
                        "tb_side": [grp(g, ccy) for g in groups], "submission_side": side(names["main"], text, ccy)})
    for key, label, kindname in (("FC", "by currency", "foreign"), ("MATURITY", "total (%s by Maturity)" % kind, "maturity")):
        if key not in inc:
            continue
        if have and names[kindname] not in have:
            if notes is not None and ("sheet", names[kindname]) not in notes:
                notes.append(("sheet", names[kindname]))
            continue
        if key == "FC":
            out.append({"label": "%s - %s" % (t["label"], label), "currency": "FC", "rule_type": "FC_COLUMNS", "derived": True,
                        "tb_groups": list(groups), "submission_side": side(names["foreign"], t.get("row_foreign") or text, "TOTAL")})
        else:
            out.append({"label": "%s - %s" % (t["label"], label), "currency": "TOTAL", "rule_type": "GROUP_RULE", "_expanded": True,
                        "derived": True, "tb_side": [grp(g, "TOTAL") for g in groups],
                        "submission_side": side(names["maturity"], t.get("row_maturity") or text, "TOTAL")})
    return out


def _rule_key(t):
    tb = []
    for x in t.get("tb_side") or []:
        if x.get("level") != "group":
            return None
        tb.append((str(x.get("bs_mapping")).lower(), x.get("currency"), x.get("sign_filter"), tuple(x.get("include_desc") or ()),
                   tuple(x.get("exclude_desc") or ()), tuple(sorted(x.get("exclude_currencies") or ())), x.get("sign", 1)))
    if not tb:
        return None
    subs = tuple((x.get("sheet"), re.sub(r"\W+", "", str(x.get("match_text")).lower())[:30], x.get("currency"), x.get("sign", 1))
                 for x in t.get("submission_side") or [])
    return (tuple(sorted(tb)), subs)


def expand_templates(templates, tb, sub, notes=None):
    """Concrete templates. ASSET_GROUP and FC_COLUMNS templates expand from the target sheets' own currency columns;
    a derived (ASSET_GROUP) rule that an explicit rule already covers is dropped, so nothing is counted twice."""
    staged = []
    for t in templates:
        if t.get("rule_type") in _KINDS:
            staged.extend(_asset_group(t, sub, notes))
        else:
            staged.append(t)
    out = []
    for t in staged:
        if t.get("rule_type") == "FC_COLUMNS":
            out.extend(_fc_expand(t, sub))
        else:
            out.append(t)
    explicit = set(k for k in (_rule_key(t) for t in out if not t.get("derived")) if k)
    return [t for t in out if not (t.get("derived") and _rule_key(t) in explicit)]


def grp(bsm, ccy, **kw):
    d = {"level": "group", "bs_mapping": bsm, "currency": ccy, "sign": 1}
    d.update(kw)
    return d


def copy_sub(s0, ccy):
    d = dict(s0)
    d["currency"] = ccy
    return [d]


# ------------------------------------------------------------------------------------------ off-balance checks
# The same basic rules as Assets / Liabilities, but the values come from the Outstanding Report pivots:
#   main sheet ('Off-Balance Sheet Accounts')      IQD column 3 or 4 / foreign column 5 or 6 / total column 2  <- Pivot 2 (LC/GTEE x currency)
#   'Off-Balance Sheet Accounts by F(oreign ...)'  currency by currency, the rest against 'Other'             <- Pivot 2
#   'Off-Balance Sheet Accounts by M(aturity)'     one check per maturity bucket column, plus the total       <- Pivot 1 (Bucket columns)

OB_SOURCE = "Outstanding Report"
_NUMWORDS = {"one": "1", "two": "2", "three": "3", "four": "4", "five": "5", "six": "6", "seven": "7", "eight": "8", "nine": "9",
             "ten": "10", "eleven": "11", "twelve": "12"}


def _ob_sheet_names(sub):
    have = sorted(set(sub.sheet)) if (sub is not None and len(sub)) else []
    main, foreign, maturity = [], None, None
    for sh in have:
        low = re.sub(r"\s+", " ", re.sub(r"[^a-z0-9 ]", " ", str(sh).lower())).strip()
        if not (low.startswith("off") and "balance" in low):
            continue
        if "matur" in low or "period" in low or "tenor" in low or low.endswith(" by m"):
            maturity = sh                                              # 'Off-Balance Sheet Accounts by M' (name cut at 31 characters)
        elif "foreign" in low or "curren" in low or low.endswith(" by f"):
            foreign = sh
        else:
            main.append(sh)
    return {"main": min(main, key=len) if main else None, "foreign": foreign, "maturity": maturity}


def _sheet_columns(sub, sheet):
    """[(column letter, header text)] of the extracted cells of a sheet, left to right."""
    out = {}
    d = sub[sub.sheet == sheet]
    for r in d.itertuples(index=False):
        m = re.match(r"^([A-Z]+)\d+$", str(r.source_cell))
        if m and m.group(1) not in out:
            out[m.group(1)] = str(getattr(r, "column_context", "") or "").split(" | ")[0].strip()      # drop the sample value the context carries
    return [(L, out[L]) for L in sorted(out, key=lambda L: (len(L), L))]


def _upper_bound_months(label):
    """Upper bound, in months, of a maturity label ('1M', '0-30 Days', 'Between Three and Six Months', 'Above 365 Days',
    'More than five years' ...). inf for an open-ended label, None when nothing can be read."""
    s = str(label).lower()
    for w, n in _NUMWORDS.items():
        s = re.sub(r"\b%s\b" % w, n, s)
    toks = [(float(a), u) for a, u in re.findall(r"(\d+(?:\.\d+)?)\s*(days?|d\b|months?|mos?\b|m\b|years?|yrs?|y\b)?", s)]
    if not toks:
        return None
    unit = None
    fixed = []
    for v, u in reversed(toks):                       # a unit written once applies to the numbers before it ("1-5Y")
        unit = u or unit
        fixed.append((v, unit))
    fixed.reverse()

    def months(v, u):
        u = (u or "m")[0]
        return v / 30.0 if u == "d" else (v * 12.0 if u == "y" else v)

    vals = [months(v, u) for v, u in fixed]
    if re.search(r"above|more than|over|greater|exceed|>|\+", s) and not re.search(r"less|under|below|<", s):
        return float("inf")
    return max(vals)


def _bucket_columns(buckets, sheet_cols, total_letter):
    """bucket label -> column letter, the header of each column, and a note when the two cannot be paired safely."""
    skip = re.compile(r"without|no maturity|undefined|unspecified|blank|not specified", re.I)
    cols = [(L, h) for L, h in sheet_cols if L != total_letter]
    open_cols = [(L, h) for L, h in cols if skip.search(h)]
    cols = [(L, h) for L, h in cols if not skip.search(h)]
    real = [b for b in buckets if b not in ("(blank)", "")]
    big = 1e9
    num = lambda v: big if v == float("inf") else v
    pair = {}
    hi_cols = [(L, _upper_bound_months(h)) for L, h in cols]
    if all(x[1] is not None for x in hi_cols):
        for b in real:
            hb = _upper_bound_months(b)
            if hb is None:
                continue
            best = None
            for L, hc in hi_cols:
                close = (hb == hc) or (hb != float("inf") and hc != float("inf") and abs(hb - hc) <= max(0.35, 0.12 * hc))
                if close and (best is None or abs(num(hb) - num(hc)) < best[0]):
                    best = (abs(num(hb) - num(hc)), L)
            if best:
                pair[b] = best[1]
    if len(pair) != len(real) or len(set(pair.values())) != len(real):
        pair = dict(zip(real, [L for L, _ in cols])) if len(real) == len(cols) else {}      # fall back to left-to-right order
    note = None
    if len(pair) != len(real):
        note = ("The maturity buckets of the Outstanding Report (%s) could not be paired with the maturity columns of the sheet (%s), "
                "so only the maturity total is checked." % (", ".join(real), ", ".join(h for _, h in cols)))
        pair = {}
    if "(blank)" in buckets and open_cols:
        pair["(blank)"] = open_cols[0][0]
    return pair, dict(cols + open_cols), note


def ob_templates(res, sub, notes):
    """Concrete rules for the off-balance sheets, from the processed Outstanding Report `res`.
    Returns (rules, pivots for the workbook)."""
    names = _ob_sheet_names(sub)
    if not any(names.values()):
        return [], None
    if res is None:
        notes.append("Off-balance sheets were extracted but no Outstanding Report is loaded, so the off-balance checks were skipped "
                     "(upload it in the Off-Balance step, then press Re-run).")
        return [], None
    try:
        from .outstanding_report import order_pivot_columns, BLANK, GRAND
    except Exception:
        notes.append("The off-balance checks need the latest outstanding_report.py in the library.")
        return [], None
    df = res["df"]
    d = df[df["_STATUS"] == "ok"].copy()
    for c in ("CATEGORY", "LC_GTEE_DESC", "BILL_CCY", "BUCKET"):
        d[c] = d[c].fillna(BLANK)
    sel = res.get("selected_category")
    d2 = d if sel in (None, "", "All") else d[d["CATEGORY"] == sel]            # Pivot 2 carries the CATEGORY filter
    p2cols = list(res["pivot2"]["flat"]["columns"])
    p1cols = list(res["pivot1"]["flat"]["columns"])
    ccys = [c for c in p2cols if c != GRAND]
    total_name = "Total Off-Balance Sheet Accounts"
    items = [("total", total_name)]
    items += [("category", c) for c in order_pivot_columns(list(d["CATEGORY"].unique()), "currency") if c != BLANK]
    items += [("lc", x) for x in order_pivot_columns(list(d["LC_GTEE_DESC"].unique()), "currency") if x != BLANK]

    def rows_of(src, kind, name):
        if kind == "category":
            return src[src["CATEGORY"] == name]
        if kind == "lc":
            return src[src["LC_GTEE_DESC"] == name]
        return src

    out = []

    def rule(label, ccy, kind, name, amount, sheet, ccy_sub, pivot, cols, letter=None):
        mark = None if (kind == "category" and pivot == 2) else ["ob", pivot, kind, name, list(cols)]
        shown = "Total" if kind == "total" else name
        side = {"match_text": name, "sheet": sheet, "currency": ccy_sub, "is_total": False, "sign": 1}
        if letter:
            side["column_letter"] = letter
        out.append({"label": "Off-balance %s - %s" % (shown, label), "currency": ccy, "rule_type": "GROUP_RULE",
                    "_expanded": True, "derived": True, "ob": True,
                    "tb_side": [{"level": "external", "label": "%s | %s" % ("Grand Total" if kind == "total" else name, label),
                                 "amount": amount, "currency": ccy, "source": OB_SOURCE, "mark": mark}],
                    "submission_side": [side]})

    # ---- main sheet and the by-currency sheet (Pivot 2)
    direct, has_iqd_col, other_col = [], False, False
    if names["foreign"]:
        sc = set(sub[sub.sheet == names["foreign"]].currency)
        has_iqd_col = "IQD" in sc
        direct = [c for c in sorted(sc) if c not in ("TOTAL", "IQD", FRX_CODE, "UNSPECIFIED") and not str(c).startswith("TOTAL_")]
        other_col = FRX_CODE in sc
    for kind, name in items:
        rows = rows_of(d if kind in ("total", "category") else d2, kind, name)
        iq = float(rows[rows["BILL_CCY"] == LOCAL_CCY]["EQUI_IQD"].sum())
        fx = float(rows[rows["BILL_CCY"] != LOCAL_CCY]["EQUI_IQD"].sum())
        tt = float(rows["EQUI_IQD"].sum())
        if names["main"]:
            rule("IQD (column 3 or 4)", "IQD", kind, name, iq, names["main"], "IQD", 2, [LOCAL_CCY])
            rule("foreign currencies (column 5 or 6)", FRX_CODE, kind, name, fx, names["main"], FRX_CODE, 2,
                 [c for c in ccys if c != LOCAL_CCY])
            rule("total (column 2)", "TOTAL", kind, name, tt, names["main"], "TOTAL", 2, [GRAND])
        if names["foreign"]:
            if has_iqd_col:
                rule("by currency - IQD", "IQD", kind, name, iq, names["foreign"], "IQD", 2, [LOCAL_CCY])
            for c in direct:
                rule("by currency - %s" % c, c, kind, name, float(rows[rows["BILL_CCY"] == c]["EQUI_IQD"].sum()),
                     names["foreign"], c, 2, [c])
            if other_col:
                rest = rows[(rows["BILL_CCY"] != LOCAL_CCY) & (~rows["BILL_CCY"].isin(direct))]
                rule("by currency - Other Foreign Currencies", FRX_CODE, kind, name, float(rest["EQUI_IQD"].sum()), names["foreign"],
                     FRX_CODE, 2, [c for c in ccys if c != LOCAL_CCY and c not in direct])

    # ---- maturity sheet (Pivot 1, Bucket columns)
    if names["maturity"]:
        cols = _sheet_columns(sub, names["maturity"])
        total_letter = next((L for L, h in cols if "total" in h.lower()), cols[0][0] if cols else None)
        buckets = [b for b in p1cols if b != GRAND]
        pair, header_of, note = _bucket_columns(buckets, cols, total_letter)
        if note:
            notes.append(note)
        for kind, name in items:
            rows = rows_of(d, kind, name)
            if total_letter:
                rule("maturity total", "TOTAL", kind, name, float(rows["EQUI_IQD"].sum()), names["maturity"], None, 1, [GRAND], total_letter)
            for b, L in pair.items():
                rule("bucket %s (column %s, %s)" % (b, L, header_of.get(L, "")), "TOTAL", kind, name,
                     float(rows[rows["BUCKET"] == b]["EQUI_IQD"].sum()), names["maturity"], None, 1, [b], L)
    return out, {"pivot1": res["pivot1"]["flat"], "pivot2": res["pivot2"]["flat"]}


# ------------------------------------------------------------------------------------------ the run

def _cells_text(components):
    by = OrderedDict()
    for c in components:
        if c.get("placeholder"):
            by.setdefault((c["submission_file"], c["sheet"]), []).append("(blank in this currency, row %s)" % c.get("row_number"))
        else:
            by.setdefault((c["submission_file"], c["sheet"]), []).append(c["source_cell"])
    return "; ".join("%s!%s" % (sh, "+".join(cells)) for (f, sh), cells in by.items())


def run_default_mapping(eng, tb, sub, templates=None, outstanding=None):
    templates = templates if templates is not None else eng.load_default_mapping()
    skipped = []
    concrete = expand_templates(templates, tb, sub, skipped)
    ob_notes = []
    ob_rules, ob_pivots = ob_templates(outstanding, sub, ob_notes)
    concrete = list(concrete) + ob_rules
    scale_of = {}
    if sub is not None and len(sub):
        for r in sub[["submission_file", "sheet", "source_cell", "scale"]].itertuples(index=False):
            scale_of[(r.submission_file, r.sheet, r.source_cell)] = float(r.scale)
    results = []
    for t in concrete:
        resolved, warnings = eng.resolve_rule_template(t, tb, sub)
        tbc = resolved.get("tb_components") or []
        sbc = resolved.get("components") or []
        tb_amt = float(sum(c["amount"] * c.get("sign", 1) for c in tbc))
        sub_amt = float(sum(c["amount"] * c.get("sign", 1) * c.get("multiplier", 1) for c in sbc))
        if t.get("_expanded") and abs(tb_amt) < 1e-9 and (not sbc or abs(sub_amt) < 1e-9):
            continue                                                  # a currency neither side has anything for
        allow = 0.0
        for c in sbc:
            sc = scale_of.get((c.get("submission_file"), c.get("sheet"), c.get("source_cell")), 1.0)
            if sc > 1:
                allow += 0.5 * sc * abs(c.get("multiplier", 1))     # figures stated in thousands carry +/- half a unit
        var = tb_amt - sub_amt
        tol = max(eng.tolerance_abs, abs(tb_amt) * eng.tolerance_pct)
        rounding = False
        if not sbc:
            status = "NOT_FOUND"
        elif abs(var) <= 1e-6:
            status = "MATCH"
        elif abs(var) <= tol:
            status = "MATCH_WITHIN_TOLERANCE"
        elif abs(var) <= allow:
            status, rounding = "MATCH_WITHIN_TOLERANCE", True
        else:
            status = "REVIEW_REQUIRED"
        notes = list(warnings)
        if rounding:
            notes.insert(0, "Variance is within the rounding of figures stated in thousands (+/- %s)." % format(allow, ",.0f"))
        results.append({
            "n": len(results) + 1, "label": t.get("label"), "rule_type": t.get("rule_type"), "currency": t.get("currency"),
            "rule_text": describe_rule(t), "tb_amount": tb_amt, "sub_amount": sub_amt, "variance": var,
            "variance_pct": (var / abs(tb_amt) * 100.0) if abs(tb_amt) > 1e-9 else None, "status": status,
            "files": sorted(set(c["submission_file"] for c in sbc)),
            "target_sheets": sorted(set(x.get("sheet") for x in t.get("submission_side", []) if x.get("sheet"))),
            "where": _cells_text(sbc), "tb_components": tbc, "components": sbc, "warnings": notes,
            "tb_marks": resolved.get("tb_marks", []), "expanded": bool(t.get("_expanded")), "derived": bool(t.get("derived")), "ob": bool(t.get("ob")), "resolved": resolved,
            "wildcard": any(str(x.get("bs_mapping") or "").endswith("*") for x in (t.get("tb_side") or []) if x.get("level") == "group"),
            "n_tb": sum(1 for c in tbc if not c.get("placeholder")), "n_sub": sum(1 for c in sbc if not c.get("placeholder")),
            "color": ""})
    palette_i = 0
    for r in results:
        if r["n_sub"]:
            r["color"] = PALETTE[palette_i % len(PALETTE)]
            palette_i += 1
    notes = ["The sheet '%s' was not extracted, so the checks that read it were skipped." % x[1] for x in skipped] + ob_notes
    return {"results": results, "summary": summarise(results), "groups": group_coverage(tb, results), "notes": notes, "ob_pivots": ob_pivots,
            "params": {"tol_abs": eng.tolerance_abs, "tol_pct": eng.tolerance_pct}}


def group_coverage(tb, results):
    """Every BS-mapping group of the Trial Balance and the rules that read it (the 'Total Assets' checks, which read
    every A- group, are not counted). A group no rule reads is listed with its amounts so it can be mapped."""
    by = defaultdict(set)
    for r in results:
        if r.get("wildcard"):
            continue
        for c in r["tb_components"]:
            if not c.get("placeholder"):
                by[str(c.get("bs_mapping")).strip()].add(r["n"])
    out = []
    if tb is None or not len(tb):
        return out
    t = tb.copy()
    t["_g"] = t.bs_mapping.astype(str).str.strip()
    for g, grp_ in t.groupby("_g", sort=True):
        local = grp_.tran_ccy == LOCAL_CCY
        out.append({"group": g if g else "(blank)", "iqd": float(grp_[local].adjusted_balance.sum()),
                    "frx": float(grp_[~local].adjusted_balance.sum()), "total": float(grp_.adjusted_balance.sum()),
                    "rules": sorted(by.get(g, [])), "covered": bool(by.get(g))})
    return out


def summarise(results):
    s = {"rules": len(results), "MATCH": 0, "MATCH_WITHIN_TOLERANCE": 0, "REVIEW_REQUIRED": 0, "NOT_FOUND": 0}
    for r in results:
        s[r["status"]] += 1
    s["abs_variance"] = float(sum(abs(r["variance"]) for r in results if r["status"] in ("REVIEW_REQUIRED", "MATCH_WITHIN_TOLERANCE")))
    return s


def public_results(run):
    """JSON for the web page (no heavy internals)."""
    out = []
    for r in run["results"]:
        out.append({k: r[k] for k in ("n", "label", "rule_type", "currency", "rule_text", "tb_amount", "sub_amount", "variance",
                                      "variance_pct", "status", "files", "target_sheets", "where", "warnings", "color",
                                      "n_tb", "n_sub", "expanded", "derived")})
    return out


# ------------------------------------------------------------------------------------------ annotated workbook

def _results_for_file(results, label, sheetnames):
    return [r for r in results if label in r["files"] or (not r["files"] and set(r["target_sheets"]) & set(sheetnames))]


def write_default_workbook(path_in, path_out, label, tb, run):
    results = run["results"]
    wb = load_workbook(path_in, keep_vba=path_in.lower().endswith(".xlsm"))
    names = set(wb.sheetnames)
    mine = _results_for_file(results, label, names)
    bold = Font(bold=True)
    grey = _fill("D9D9D9")

    # -- highlight the matched submission cells, with a comment naming the rule and its variance
    notes = defaultdict(list)
    for r in mine:
        if not r["color"]:
            continue
        for c in r["components"]:
            if c.get("submission_file") != label or c.get("placeholder") or c["sheet"] not in names:
                continue
            key = (c["sheet"], c["source_cell"])
            if not notes[key]:
                wb[c["sheet"]][c["source_cell"]].fill = _fill(r["color"])
            notes[key].append("Rule %d: %s | TB %s | variance %s" % (r["n"], r["label"][:70], format(r["tb_amount"], ",.0f"),
                                                                   format(r["variance"], ",.0f")))
    for (sheet, cell), lines in list(notes.items())[:3000]:
        cm = Comment("Default mapping:\n" + "\n".join(lines[:4]), "Default Mapping")
        cm.width, cm.height = 380, 40 + 18 * min(len(lines), 4)
        wb[sheet][cell].comment = cm

    # -- TB pivot (BS-mapping level) and TB accounts, coloured the same way
    layout = build_layout(tb, include_accounts=False, min_value=0.0)
    keys = layout["keys"]
    piv_pos = {}
    for i, (level, lab, vec, _n) in enumerate(layout["rows"]):
        if level == "BS Mapping":
            piv_pos[lab] = i + 2
    acct_pos = dict(((b, str(a), str(d)), i + 2) for i, (b, a, d, _v) in enumerate(layout["acct_rows"]))
    wp = wb.create_sheet(_unique_name(wb, "TB Pivot"))
    wp.append(["Level", "BS Mapping"] + [_col_title(k) for k in keys])
    for c in wp[1]:
        c.font, c.fill = bold, grey
    for level, lab, vec, _n in layout["rows"]:
        wp.append([level, lab] + [vec[k] for k in keys])
    wa = wb.create_sheet(_unique_name(wb, "TB Accounts"))
    wa.append(["BS Mapping", "Account", "Account Desc"] + [_col_title(k) for k in keys])
    for c in wa[1]:
        c.font, c.fill = bold, grey
    for b, a, d, vec in layout["acct_rows"]:
        wa.append([b, a, d] + [vec[k] for k in keys])
    for ws_, first in ((wp, 3), (wa, 4)):
        for row in ws_.iter_rows(min_row=2, min_col=first):
            for cell in row:
                cell.number_format = "#,##0.00;[Red]-#,##0.00"
    for r in mine:
        if not r["color"]:
            continue
        for mk in r["tb_marks"]:
            if mk[0] == "pivot" and mk[1] in piv_pos and mk[2] in keys:
                wp.cell(row=piv_pos[mk[1]], column=3 + keys.index(mk[2])).fill = _fill(r["color"])
            elif mk[0] == "acct":
                row_i = acct_pos.get((mk[1], mk[2], mk[3]))
                if row_i and mk[4] in keys:
                    wa.cell(row=row_i, column=4 + keys.index(mk[4])).fill = _fill(r["color"])
    wp.freeze_panes = "C2"
    wa.freeze_panes = "D2"
    _write_ob_pivots(wb, run, mine, bold, grey)
    wp.column_dimensions["B"].width = 36
    wa.column_dimensions["A"].width = 18
    wa.column_dimensions["C"].width = 38

    # -- the report: every rule, how it works, TB vs submission, variance
    wr = wb.create_sheet(_unique_name(wb, "Default Mapping Report"))
    wr["A1"] = "Default mapping - results for %s" % label
    wr["A1"].font = Font(bold=True, size=13)
    wr["A2"] = ("Tolerance: %s absolute or %s%% of the TB amount. Variance = TB amount - submission amount. "
                "Matched cells and the TB values they were compared with share one colour." % (
                    run["params"]["tol_abs"], format(run["params"]["tol_pct"] * 100, "g")))
    head = ["#", "Colour", "Rule", "How the rule works", "Currency", "TB amount", "Submission sheet", "Submission cells",
            "Submission amount", "Variance (TB - submission)", "Variance %", "Status", "Notes"]
    wr.append([])
    wr.append(head)
    for c in wr[4]:
        c.font, c.fill = bold, grey
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for r in mine:
        sheets = ", ".join(sorted(set(c["sheet"] for c in r["components"]))) or ", ".join(r["target_sheets"])
        wr.append([r["n"], "", r["label"], r["rule_text"], r["currency"], r["tb_amount"], sheets, r["where"] or "(not found)",
                   r["sub_amount"], r["variance"], r["variance_pct"], STATUS_TEXT[r["status"]], " | ".join(r["warnings"][:3])])
        ri = wr.max_row
        if r["color"]:
            wr.cell(row=ri, column=2).fill = _fill(r["color"])
        wr.cell(row=ri, column=12).fill = _fill(STATUS_HEX[r["status"]])
        for col in (6, 9, 10):
            wr.cell(row=ri, column=col).number_format = "#,##0.00;[Red]-#,##0.00"
        wr.cell(row=ri, column=11).number_format = '0.0000"%"'
        first = next((c for c in r["components"] if c.get("submission_file") == label and not c.get("placeholder")), None)
        if first and first["sheet"] in names:
            cell = wr.cell(row=ri, column=8)
            cell.hyperlink = Hyperlink(ref=cell.coordinate, location="%s!%s" % (_sheet_ref(first["sheet"]), first["source_cell"]),
                                       display=r["where"])
            cell.font = Font(color="0563C1", underline="single")
        for col in range(1, len(head) + 1):
            wr.cell(row=ri, column=col).alignment = Alignment(wrap_text=True, vertical="top")
    for letter, w in zip("ABCDEFGHIJKLM", [5, 8, 46, 70, 9, 18, 20, 30, 18, 20, 11, 22, 50]):
        wr.column_dimensions[letter].width = w
    wr.freeze_panes = "D5"
    wr.auto_filter.ref = "A4:M%d" % max(wr.max_row, 5)

    # -- every Trial Balance account behind every rule
    wd = wb.create_sheet(_unique_name(wb, "Default Mapping TB Detail"))
    wd.append(["Rule #", "Rule", "BS mapping", "Account", "Account description", "Currency", "Amount", "Currencies behind it"])
    for c in wd[1]:
        c.font, c.fill = bold, grey
    for r in mine:
        for c in r["tb_components"]:
            wd.append([r["n"], r["label"][:80], c.get("bs_mapping"), c.get("account"), c.get("account_desc"), c.get("currency"),
                       c.get("amount", 0) * c.get("sign", 1),
                       ", ".join("%s %s" % (k, format(v, ",.0f")) for k, v in (c.get("breakdown") or {}).items())])
            wd.cell(row=wd.max_row, column=7).number_format = "#,##0.00;[Red]-#,##0.00"
    for letter, w in zip("ABCDEFGH", [7, 50, 14, 12, 40, 10, 18, 40]):
        wd.column_dimensions[letter].width = w
    wd.freeze_panes = "A2"

    # -- every TB group and the rules that read it
    wg = wb.create_sheet(_unique_name(wb, "Default Mapping Groups"))
    wg.append(["BS mapping group", "IQD", "Foreign currencies", "Total", "Rules (#) that read it", "Covered?"])
    for c in wg[1]:
        c.font, c.fill = bold, grey
    for g in run.get("groups") or []:
        wg.append([g["group"], g["iqd"], g["frx"], g["total"], ", ".join(str(n) for n in g["rules"]),
                   "Yes" if g["covered"] else "NO RULE"])
        for col in (2, 3, 4):
            wg.cell(row=wg.max_row, column=col).number_format = "#,##0.00;[Red]-#,##0.00"
        if not g["covered"]:
            wg.cell(row=wg.max_row, column=6).fill = _fill(STATUS_HEX["NOT_FOUND"])
    for letter, w in zip("ABCDEF", [24, 20, 20, 20, 40, 12]):
        wg.column_dimensions[letter].width = w
    wg.freeze_panes = "A2"

    # -- summary
    ws = wb.create_sheet(_unique_name(wb, "Default Mapping Summary"))
    sm = summarise(mine)
    for a, b in [("Submission file", label), ("Generated", datetime.datetime.now().strftime("%Y-%m-%d %H:%M")),
                 ("Rules that apply to this file", sm["rules"]), ("Match", sm["MATCH"]),
                 ("Match within tolerance", sm["MATCH_WITHIN_TOLERANCE"]), ("Variance - review", sm["REVIEW_REQUIRED"]),
                 ("Not found", sm["NOT_FOUND"]), ("Total absolute variance (rules with a variance)", sm["abs_variance"])]:
        ws.append([a, b])
    for i in range(1, ws.max_row + 1):
        ws.cell(row=i, column=1).font = bold
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 40
    wb.active = wb.sheetnames.index(wr.title)
    for sh in wb.worksheets:
        sh.sheet_view.tabSelected = sh.title == wr.title
    wb.save(path_out)
    return path_out


def _write_ob_pivots(wb, run, mine, bold, grey):
    """OB Pivot 1 / OB Pivot 2 (the Outstanding Report) with the cells the off-balance rules used coloured."""
    pv = run.get("ob_pivots")
    if not pv or not any(m[0] == "ob" for r in mine for m in r["tb_marks"]):
        return
    sheets = {}
    for n in (1, 2):
        flat = pv["pivot%d" % n]
        ws = wb.create_sheet(_unique_name(wb, "OB Pivot %d" % n))
        nl = len(flat["label_cols"])
        ws.append(list(flat["label_cols"]) + list(flat["columns"]))
        for c in ws[1]:
            c.font, c.fill = bold, grey
        for r in flat["rows"]:
            ws.append(list(r["labels"]) + list(r["values"]) + [r["total"]])
            for col in range(nl + 1, nl + 2 + len(r["values"])):
                ws.cell(row=ws.max_row, column=col).number_format = "#,##0.00;[Red]-#,##0.00"
        ws.freeze_panes = ws.cell(row=2, column=nl + 1)
        ws.column_dimensions["A"].width = 26
        sheets[n] = (ws, flat, nl)
    for r in mine:
        if not r["color"]:
            continue
        for mk in r["tb_marks"]:
            if mk[0] != "ob":
                continue
            _, n, kind, name, colnames = mk
            ws, flat, nl = sheets[n]
            for i, row in enumerate(flat["rows"]):
                lv, lab = row["level"], row["labels"]
                hit = (kind == "total" and lv == "Grand Total") or \
                      (kind == "category" and n == 1 and lv == "Category" and lab[0] == name) or \
                      (kind == "lc" and lv == "LC/GTEE" and (lab[1] if n == 1 else lab[0]) == name)
                if not hit:
                    continue
                for cn in colnames:
                    if cn in flat["columns"]:
                        ws.cell(row=i + 2, column=nl + 1 + flat["columns"].index(cn)).fill = _fill(r["color"])


def build_zip(entries):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, path in entries:
            z.write(path, name)
    return buf.getvalue()


# ------------------------------------------------------------------------------------------ PDF

def _money(v):
    return "" if v is None else format(v, ",.2f")


def build_pdf(run, files=None, tb_name="", suggestions=None, approved=None, manual_matches=None):
    results = run["results"]
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    pdf = SimplePDF(landscape=True, title="Bahrain-Iraq Recon Studio - Default Mapping", footer="Generated " + now)
    pdf.heading("Default Mapping - results", 17)
    pdf.paragraph("Trial Balance: %s    |    Tolerance: %s absolute or %s%% of the TB amount    |    Variance = TB amount - submission amount"
                  % (tb_name or "(uploaded)", run["params"]["tol_abs"], format(run["params"]["tol_pct"] * 100, "g")), 8.5)
    pdf.paragraph("Amounts are in the same units as the Trial Balance (figures a sheet states in thousands are converted). "
                  "Text outside the Latin alphabet, such as Arabic, cannot be drawn in this PDF and appears as ? - "
                  "translate the submissions first if you need it readable here.", 7.5, color=(0.4, 0.4, 0.4))
    s = run["summary"]
    pdf.heading("1. Summary", 12)
    pdf.table(["Rules applied", "Match", "Match (within tolerance)", "Variance - review", "Not found", "Total absolute variance"],
              [[s["rules"], s["MATCH"], s["MATCH_WITHIN_TOLERANCE"], s["REVIEW_REQUIRED"], s["NOT_FOUND"], _money(s["abs_variance"])]],
              [1, 1, 1.4, 1.2, 1, 1.6], size=9, aligns=["r"] * 6)
    if files:
        rows = []
        for f in files:
            mine = _results_for_file(results, f["file"], f.get("sheets") or [])
            sm = summarise(mine)
            rows.append([f["file"], f.get("type_label", ""), sm["rules"], sm["MATCH"] + sm["MATCH_WITHIN_TOLERANCE"],
                         sm["REVIEW_REQUIRED"], sm["NOT_FOUND"]])
        pdf.table(["Submission file", "Type", "Rules", "Matched", "Variance - review", "Not found"], rows, [5, 2.5, 1, 1, 1.4, 1],
                  size=8, aligns=["l", "l", "r", "r", "r", "r"])

    pdf.heading("2. Results by rule", 12)
    rows, fills = [], []
    for r in results:
        pct = "" if r["variance_pct"] is None else format(r["variance_pct"], ".4f") + "%"
        rows.append([r["n"], r["label"], r["rule_text"], r["currency"], _money(r["tb_amount"]), _money(r["sub_amount"]),
                     _money(r["variance"]), pct, STATUS_TEXT[r["status"]], r["where"] or "(not found)"])
        fills.append({8: STATUS_RGB[r["status"]]})
    pdf.table(["#", "Rule", "How the rule works", "Ccy", "TB amount", "Submission amount", "Variance", "Var %", "Status", "Where in the submission"],
              rows, [0.5, 3.4, 4.6, 0.8, 1.6, 1.6, 1.5, 1.0, 1.5, 2.6], size=6.2,
              aligns=["l", "l", "l", "l", "r", "r", "r", "r", "l", "l"], fills=fills)

    groups = run.get("groups") or []
    if groups:
        pdf.heading("3. Trial Balance groups and the rules that read them", 12)
        un = [g for g in groups if not g["covered"] and (abs(g["total"]) > 1e-6)]
        pdf.paragraph("%d of %d BS-mapping groups are read by at least one rule. %s" % (
            len(groups) - len(un), len(groups),
            ("Groups no rule reads yet (they still need a mapping): " + ", ".join(g["group"] for g in un)) if un else
            "Every group with a balance is covered."), 8, color=(0.1, 0.1, 0.1))
        pdf.table(["BS mapping group", "IQD", "Foreign currencies", "Total", "Rules"],
                  [[g["group"], _money(g["iqd"]), _money(g["frx"]), _money(g["total"]),
                    ", ".join("#%d" % n for n in g["rules"]) if g["rules"] else "NO RULE"] for g in groups],
                  [2, 1.6, 1.8, 1.6, 4], size=6.6, aligns=["l", "r", "r", "r", "l"],
                  fills=[{} if g["covered"] else {4: STATUS_RGB["NOT_FOUND"]} for g in groups])
    pdf.heading("4. Rule detail - Trial Balance side and submission side", 12)
    for r in results:
        pdf.ensure(70)
        pdf.paragraph("#%d  %s" % (r["n"], r["label"]), 9.5, True, gap=1)
        pdf.paragraph("Status: %s   |   TB %s   vs   submission %s   =   variance %s%s" % (
            STATUS_TEXT[r["status"]], _money(r["tb_amount"]), _money(r["sub_amount"]), _money(r["variance"]),
            "" if r["variance_pct"] is None else " (%s%%)" % format(r["variance_pct"], ".4f")), 8, color=(0.1, 0.1, 0.1), gap=1)
        pdf.paragraph("Rule: " + r["rule_text"], 7.5, color=(0.3, 0.3, 0.3), gap=3)
        tbr = []
        for c in r["tb_components"]:
            extra = ""
            if c.get("breakdown"):
                extra = "  (" + ", ".join("%s %s" % (k, format(v, ",.0f")) for k, v in c["breakdown"].items()) + ")"
            tbr.append([c.get("bs_mapping"), c.get("account"), str(c.get("account_desc")) + extra, c.get("currency"),
                        _money(c.get("amount", 0) * c.get("sign", 1))])
        tbr.append(["Total", "", "", "", _money(r["tb_amount"])])
        pdf.table(["TB - BS group", "Account", "Description", "Ccy", "Amount"], tbr, [1.3, 1, 5, 0.8, 1.6], size=6.4,
                  aligns=["l", "l", "l", "l", "r"], bold_rows=[len(tbr) - 1])
        sbr = []
        for c in r["components"]:
            sbr.append([c.get("submission_file"), c.get("sheet"), c.get("source_cell"), c.get("line_description"), c.get("currency"),
                        _money(c.get("amount")), _money(c.get("amount", 0) * c.get("sign", 1) * c.get("multiplier", 1))])
        if not sbr:
            sbr.append(["(not found)", "", "", "; ".join(r["warnings"][:2]), "", "", ""])
        sbr.append(["Total", "", "", "", "", "", _money(r["sub_amount"])])
        pdf.table(["Submission file", "Sheet", "Cell", "Line", "Ccy", "Value in sheet", "Counted"], sbr,
                  [2.2, 1.6, 0.8, 4.2, 0.7, 1.4, 1.4], size=6.4, aligns=["l", "l", "l", "l", "l", "r", "r"], bold_rows=[len(sbr) - 1])
        if r["warnings"] and r["n_sub"]:
            pdf.paragraph("Notes: " + " | ".join(r["warnings"][:4]), 6.8, color=(0.45, 0.35, 0.05), gap=6)

    if suggestions:
        pdf.heading("5. Auto-suggested matches (Mapping Studio)", 12)
        rows = []
        for i, sg in enumerate(suggestions):
            ok = bool(approved[i]) if approved and i < len(approved) else False
            rows.append([sg.get("bs_mapping"), sg.get("currency"), sg.get("rule_type"), _money(sg.get("tb_amount")),
                         _money(sg.get("suggested_submission_amount")), _money(sg.get("difference")),
                         "%d%%" % round(100 * sg.get("confidence", 0)), "Approved" if ok else "Not approved"])
        pdf.table(["BS mapping / group", "Ccy", "Rule type", "TB amount", "Submission amount", "Difference", "Confidence", "Decision"],
                  rows, [4, 0.8, 2.2, 1.6, 1.8, 1.5, 1, 1.3], size=6.4, aligns=["l", "l", "l", "r", "r", "r", "r", "l"])

    if manual_matches:
        pdf.heading("6. Manual and imported matches (Mapping Studio)", 12)
        rows = []
        for m in manual_matches:
            tb = sum(c.get("amount", 0) * c.get("sign", 1) for c in m.get("tb_components", []))
            sb = sum(c.get("amount", 0) * c.get("sign", 1) * c.get("multiplier", 1) for c in m.get("components", []))
            rows.append([m.get("label") or m.get("bs_mapping"), m.get("currency"), m.get("source", "MANUAL"),
                         len(m.get("tb_components", [])), len(m.get("components", [])), _money(tb), _money(sb), _money(tb - sb)])
        pdf.table(["Match", "Ccy", "Source", "TB items", "Submission items", "TB total", "Submission total", "Variance"], rows,
                  [5, 0.8, 1.2, 1, 1.3, 1.6, 1.8, 1.5], size=6.6, aligns=["l", "l", "l", "r", "r", "r", "r", "r"])
    return pdf.output()
