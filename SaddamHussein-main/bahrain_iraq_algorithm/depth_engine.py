"""Depth engine: take the numbers of ANY workbook (the source, the sheets you mark) and look for every one of them in the
numbers of other workbooks (the targets) - then write annotated copies and build a mapping from what was found.

* direct search   - each source value against every non-zero cell of the targets: face value, x1,000, x1,000,000 and the
                    reverse, equal-and-opposite sign (x -1), rounding to the digits the sheet shows (same engine as the
                    Trial Balance depth search: WorkbookIndex).
* bucket search   - a value no single cell holds is looked for as a COMBINATION of cells that add up to it (a subset-sum /
                    knapsack search inside one column or one row of a sheet; meet-in-the-middle, windows over big columns).
* outputs         - an annotated copy of every target (matched cells coloured, a comment on each, Source Values, Matching Report,
                    Combinations and Unmatched sheets) and of the source (found / not found).
* mapping         - what was found, as a portable mapping (labels, never cell addresses) and as default-mapping rules.

Python 3.9 compatible.
"""
import datetime
import time
from collections import OrderedDict, defaultdict

import numpy as np
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

from .depth_search import PALETTE, WorkbookIndex, _decimals, _fill, _sheet_ref, _unique_name, build_zip, currency_check
from .engine import FORMULA_REF_PATTERN, FRX_CODE, text_similarity

DEFAULT_PARAMS = {"tol_abs": 1.0, "allow_scale": True, "allow_sign": True, "min_value": 1.0, "min_scaled": 1000.0,
                  "source_scale": "auto",            # "auto" = the unit the source sheet declares; or 1 / 1000 / 1000000
                  "source_label": "nearest",         # the row label of a source number: the text nearest on its left, or the longest one
                  "max_hits": 25, "bucket": True, "bucket_cells": 20, "bucket_max_k": 8, "bucket_scope": "both",
                  "max_solutions": 3, "time_limit": 90.0,
                  "bucket_mode": "both",             # "run" = neighbouring cells only, "subset" = knapsack only, "both" = runs first
                  "max_chance": 0.2,                 # a combination is only a strong match when fewer coincidences than this are expected in the whole search
                  "max_weak": 5.0}                   # windows where even a hit would be a coincidence (more expected than this) are not searched
SCALES = (1.0, 1000.0, 1000000.0, 0.001, 0.000001)
STATUS_TEXT = {"DIRECT": "Found", "COMBINATION": "Found as a combination", "POSSIBLE": "Possible combination (could be a coincidence)",
               "NOT_FOUND": "Not found", "SKIPPED": "Not searched (time limit)"}


# ------------------------------------------------------------------------------------------ reading

def sheet_summaries(idx):
    """[{sheet, cells, hidden, scale}] - what the user marks."""
    counts = np.bincount(idx.sidx, minlength=len(idx.sheets)) if idx.n_cells else np.zeros(len(idx.sheets), dtype=int)
    return [{"sheet": sh.name, "cells": int(counts[k]), "hidden": bool(sh.hidden), "scale": float(sh.scale_hint)}
            for k, sh in enumerate(idx.sheets)]


def _nearest_label(sh, row, col):
    """The text cell nearest to the left of a number (skipping short codes and formula references)."""
    for c in range(col - 1, 0, -1):
        s = sh.strs.get((row, c))
        if s and len(s) >= 3 and not FORMULA_REF_PATTERN.search(s):
            return s[:140]
    return ""


def source_values(idx, sheets, params):
    """Every non-zero number of the marked sheets, with its row label, column header, column currency and unit."""
    p = dict(DEFAULT_PARAMS)
    p.update(params or {})
    mode = p.get("source_scale", "auto")
    want = set(sheets or [])
    out = []
    for i in range(idx.n_cells):
        sh = idx.sheets[int(idx.sidx[i])]
        if want and sh.name not in want:
            continue
        row, col = int(idx.rows[i]), int(idx.cols[i])
        raw = float(idx.raw[i])
        scale = float(sh.column_scale(col)) if str(mode) == "auto" else float(mode)
        amount = raw * scale
        if abs(amount) < float(p["min_value"]):
            continue
        out.append({"id": len(out), "sheet": sh.name, "cell": "%s%d" % (get_column_letter(col), row), "row": row, "col": col,
                    "row_label": (_nearest_label(sh, row, col) or sh.label(row, col)) if p.get("source_label", "nearest") == "nearest" else sh.label(row, col),
                    "column_header": sh.header(col), "currency": sh.column_ccy(col) or "",
                    "raw": raw, "scale": scale, "amount": amount})
    out.sort(key=lambda v: (v["sheet"], v["row"], v["col"]))
    for n, v in enumerate(out):
        v["id"] = n
    return out


# ------------------------------------------------------------------------------------------ direct search

def _direct_hits(v, targets, p):
    """Every cell of every target that holds this value (any unit / sign allowed by the settings), best first."""
    absT = abs(v["amount"])
    scales = [1.0] + (list(SCALES[1:]) if p["allow_scale"] else [])
    tol_abs = float(p["tol_abs"])
    hits = []
    for tg in targets:
        idx = tg["idx"]
        wanted = tg.get("sheets")
        cands = {}
        for s in scales:
            if s > 1.0 and absT < max(float(p["min_scaled"]), 5.0 * s):
                continue
            win = tol_abs if s <= 1.0 else max(tol_abs, 0.5 * s)
            for i in idx.find(absT, s, win):
                sh = idx.sheets[int(idx.sidx[i])]
                if wanted and sh.name not in wanted:
                    continue
                raw = float(idx.raw[i])
                tol_c = tol_abs if s <= 1.0 else max(tol_abs, 0.5 * s * (10.0 ** (-_decimals(raw))))
                diff = abs(abs(raw) * s - absT)
                if diff > tol_c + 1e-9:
                    continue
                opposite = (raw > 0) != (v["amount"] > 0)
                if opposite and not p["allow_sign"]:
                    continue
                prev = cands.get(int(i))
                if prev is None or (abs(np.log10(s)), diff) < (abs(np.log10(prev[0])), prev[1]):
                    cands[int(i)] = (s, diff, opposite)
        for i, (s, diff, opposite) in cands.items():
            sh = idx.sheets[int(idx.sidx[i])]
            row, col = int(idx.rows[i]), int(idx.cols[i])
            rlabel = sh.label(row, col)
            ccy = sh.column_ccy(col)
            chk = currency_check(v["currency"], ccy) if v["currency"] else "UNKNOWN"
            sim = text_similarity(rlabel, v["row_label"]) if (rlabel and v["row_label"]) else 0.0
            scale_pen = 0 if (s == 1.0 or s == sh.column_scale(col)) else 1
            hits.append(((chk == "CONFLICT", scale_pen, round(diff, 6), -sim), {
                "file": tg["label"], "sheet": sh.name, "cell": "%s%d" % (get_column_letter(col), row), "row": row, "col": col,
                "value": float(idx.raw[i]), "matched_value": float(abs(idx.raw[i]) * s), "scale": s, "opposite_sign": bool(opposite),
                "diff": float(diff), "method": "DIRECT", "row_label": rlabel, "column_header": sh.header(col),
                "column_currency": ccy or "", "currency_check": chk, "name_similarity": round(float(sim), 3)}))
    hits.sort(key=lambda x: x[0])
    return [h for _, h in hits[: int(p["max_hits"])]]


# ------------------------------------------------------------------------------------------ bucket (knapsack) search

def _enumerate(values):
    sums = np.zeros(1)
    masks = np.zeros(1, dtype=np.int64)
    for j, x in enumerate(values):
        sums = np.concatenate([sums, sums + x])
        masks = np.concatenate([masks, masks | (1 << j)])
    return sums, masks


def _popcount(masks):
    pc = np.zeros(len(masks), dtype=np.int16)
    m = masks.copy()
    while m.any():
        pc += (m & 1).astype(np.int16)
        m >>= 1
    return pc


def subset_search(vals, target, window, max_k, max_sol):
    """Subsets (2..max_k items) of `vals` whose sum is within `window` of `target` - meet in the middle. Returns
    [(size, |sum - target|, [indices])], smallest first."""
    n = len(vals)
    if n < 2:
        return []
    h = n // 2
    sa, ma = _enumerate(vals[:h])
    sb, mb = _enumerate(vals[h:])
    order = np.argsort(sb, kind="stable")
    sbs, mbs = sb[order], mb[order]
    pa, pb = _popcount(ma), _popcount(mbs)
    lo = np.searchsorted(sbs, target - sa - window, side="left")
    hi = np.searchsorted(sbs, target - sa + window, side="right")
    sols = []
    for ia in np.nonzero(hi > lo)[0][:4000]:
        for jb in range(int(lo[ia]), min(int(hi[ia]), int(lo[ia]) + 50)):
            k = int(pa[ia]) + int(pb[jb])
            if k < 2 or k > max_k:
                continue
            idxs = [b for b in range(h) if (int(ma[ia]) >> b) & 1] + [h + b for b in range(n - h) if (int(mbs[jb]) >> b) & 1]
            sols.append((k, abs(float(sum(vals[i] for i in idxs)) - target), idxs))
        if len(sols) > max_sol * 40:
            break
    sols.sort(key=lambda s: (s[0], s[1]))
    return sols[:max_sol]


def run_search(vals, target, window, max_k, max_sol):
    """Runs of 2..max_k NEIGHBOURING cells (the children under a subtotal) whose sum is within `window` of `target`."""
    n = len(vals)
    out = []
    for i in range(n):
        tot = 0.0
        for j in range(i, min(n, i + max_k)):
            tot += float(vals[j])
            if j > i and abs(tot - target) <= window:
                out.append((j - i + 1, abs(tot - target), list(range(i, j + 1))))
    out.sort(key=lambda s: (s[0], s[1]))
    return out[:max_sol]


def chance(vals, target, window, n_candidates):
    """Expected number of combinations that match BY COINCIDENCE: how many were tried x the chance that one sum lands within the
    window of the target (normal approximation of the sums of random subsets of these cells)."""
    v = np.asarray(vals, dtype=float)
    if not len(v):
        return 1.0
    mu = float(v.sum()) / 2.0
    sd = float(np.sqrt((v * v).sum()) / 2.0) or 1.0
    pdf = float(np.exp(-0.5 * ((target - mu) / sd) ** 2) / (sd * np.sqrt(2.0 * np.pi)))
    return float(n_candidates * 2.0 * window * pdf)


def _pools(tg, scope):
    """Cells of one column / one row of one sheet, in sheet order (cached on the target)."""
    if tg.get("_pools") is None:
        idx = tg["idx"]
        wanted = tg.get("sheets")
        cols, rows = defaultdict(list), defaultdict(list)
        for i in range(idx.n_cells):
            sh = idx.sheets[int(idx.sidx[i])]
            if wanted and sh.name not in wanted:
                continue
            row, col = int(idx.rows[i]), int(idx.cols[i])
            cell = (row, col, float(idx.raw[i]), float(sh.column_scale(col)), sh)
            cols[(sh.name, col)].append(cell)
            rows[(sh.name, row)].append(cell)
        pools = []
        for (name, col), cells in cols.items():
            pools.append(("column", name, "%s" % get_column_letter(col), sorted(cells, key=lambda c: c[0])))
        for (name, row), cells in rows.items():
            pools.append(("row", name, str(row), sorted(cells, key=lambda c: c[1])))
        tg["_pools"] = pools
    return [x for x in tg["_pools"] if scope == "both" or x[0] == scope]


def _plan(v, targets, p, tier, mults):
    """Every (pool window, unit, sign) this tier would search for the value, each with its expected number of coincidences."""
    absT = abs(v["amount"])
    tol_abs = float(p["tol_abs"])
    cap = max(4, min(40, int(p["bucket_cells"])))
    max_k = max(2, int(p["bucket_max_k"]))
    max_weak = float(p.get("max_weak", 5.0))
    units = []
    for m in mults:
        for tg in targets:
            for kind, sheet, key, cells in _pools(tg, p["bucket_scope"]):
                if len(cells) < 2:
                    continue
                if v["currency"]:
                    # a USD number is never built from IQD cells, nor from a "Total" column (it mixes every currency)
                    def _ok(c, src=v["currency"]):
                        cc = c[4].column_ccy(c[1])
                        return currency_check(src, cc) != "CONFLICT" and not (cc == "TOTAL" and src != "TOTAL")
                    cells = [c for c in cells if _ok(c)]
                    if len(cells) < 2:
                        continue
                if m != 1.0 and all(c[3] > 1.0 for c in cells):
                    continue                                             # the sheet declares its own unit: no other unit is guessed
                norm = np.array([c[2] * c[3] * m for c in cells], dtype=float)
                unit = max([c[3] * m * 10.0 ** (-_decimals(c[2])) for c in cells] + [0.0])
                window = max(tol_abs, 0.5 * unit * max_k) if unit > 1 else tol_abs
                keep = list(range(len(cells)))
                if np.all(norm >= 0):                                    # only cells that are not bigger than the value can be in a sum
                    keep = [i for i in keep if norm[i] <= absT + window]
                    if len(keep) < 2 or norm[keep].sum() < absT - window:
                        continue
                sub_cells = [cells[i] for i in keep]
                sub_vals = norm[keep]
                if tier == "run":
                    starts, width = [0], len(sub_cells)
                else:
                    width = cap
                    starts = [0] if len(sub_cells) <= cap else list(range(0, len(sub_cells) - cap // 2, max(1, cap // 2)))
                ex_w = window                                           # the widest tolerance a hit can be accepted with
                for st in starts:
                    seg = sub_cells[st: st + width]
                    vals = sub_vals[st: st + width]
                    if len(vals) < 2:
                        continue
                    n_cand = float(len(vals) * max(1, max_k - 1)) if tier == "run" else float(2 ** len(vals))
                    for target in ([v["amount"], -v["amount"]] if p["allow_sign"] else [v["amount"]]):
                        e = chance(vals, target, ex_w, n_cand)
                        if e > max_weak:
                            continue                                    # so many candidates that any hit would be a coincidence: not searched
                        units.append({"tier": tier, "m": m, "tg": tg, "kind": kind, "sheet": sheet, "key": key, "seg": seg, "vals": vals,
                                      "window": window, "unit": unit, "target": target, "e": e})
    return units


def _combo_search(v, targets, p, deadline):
    """The value as a COMBINATION of cells of one column / row. Two tiers:
    1. RUNS - neighbouring cells (the children under a subtotal): few candidates;
    2. KNAPSACK - any cells (meet in the middle).
    Before searching, every window that will be looked at is listed with its expected number of COINCIDENTAL matches; the sum over
    the whole search is the coincidence estimate of any hit. A hit is a *strong* match only when that estimate is below
    `max_chance`; otherwise it is reported as POSSIBLE. Returns (combinations, strong?)."""
    tol_abs = float(p["tol_abs"])
    max_k = max(2, int(p["bucket_max_k"]))
    max_sol = int(p["max_solutions"])
    max_chance = float(p.get("max_chance", 0.2))
    mode = p.get("bucket_mode", "both")
    mults = [1.0] + (list(SCALES[1:]) if p["allow_scale"] else [])
    strong, weak = [], []
    searched_e = 0.0
    for tier in [t for t in ("run", "subset") if mode in ("both", t)]:
        units = _plan(v, targets, p, tier, mults)
        searched_e += sum(u["e"] for u in units)                        # everything looked at so far counts towards the coincidence estimate
        finder = run_search if tier == "run" else subset_search
        for u in units:
            if time.time() > deadline:
                break
            for k, d, idxs in finder(u["vals"], u["target"], u["window"], max_k, max_sol):
                chosen = [u["seg"][i] for i in idxs]
                m = u["m"]
                exact = max(tol_abs, 0.5 * sum(c[3] * m * 10.0 ** (-_decimals(c[2])) for c in chosen if c[3] * m > 1))
                if d > exact + 1e-9:
                    continue
                rec = {"method": "COMBINATION", "tier": "neighbouring cells" if tier == "run" else "knapsack", "file": u["tg"]["label"],
                       "sheet": u["sheet"], "pool": u["kind"], "pool_key": u["key"], "multiplier": m, "opposite_sign": u["target"] != v["amount"],
                       "diff": float(d), "k": k, "chance": round(searched_e, 4), "strong": searched_e <= max_chance,
                       "cells": [{"cell": "%s%d" % (get_column_letter(c[1]), c[0]), "row": c[0], "col": c[1], "value": c[2], "scale": c[3] * m,
                                  "row_label": c[4].label(c[0], c[1]), "column_header": c[4].header(c[1]),
                                  "column_currency": c[4].column_ccy(c[1]) or ""} for c in chosen],
                       "matched_value": float(sum(c[2] * c[3] * m for c in chosen))}
                (strong if rec["strong"] else weak).append(rec)
            if len(strong) >= max_sol:
                break
        if strong:
            break
    strong.sort(key=lambda c: (c["k"], c["diff"], c["chance"]))
    if strong:
        return strong[:max_sol], True
    weak.sort(key=lambda c: (c["k"], c["diff"]))
    return weak[:max_sol], False


# ------------------------------------------------------------------------------------------ the run

def run_engine(source, targets, params=None):
    """source: {"label", "idx", "sheets"}; targets: [{"label", "idx", "sheets": None | [names]}]. Every distinct source value is
    searched (equal amounts are searched once and shared); returns the groups, statistics and notes."""
    p = dict(DEFAULT_PARAMS)
    p.update(dict((k, v) for k, v in (params or {}).items() if v is not None))
    t0 = time.time()
    deadline = t0 + float(p["time_limit"])
    values = source_values(source["idx"], source.get("sheets"), p)
    for tg in targets:
        tg["_pools"] = None
        tg["sheets"] = set(tg["sheets"]) if tg.get("sheets") else None
    groups = OrderedDict()
    seen = {}
    for v in values:
        k = round(abs(v["amount"]), 2)
        if k in seen:
            groups[seen[k]]["members"].append(v["id"])
            v["group"] = seen[k]
        else:
            seen[k] = v["id"]
            v["group"] = v["id"]
            groups[v["id"]] = {"id": v["id"], "members": [v["id"]], "hits": [], "combos": [], "status": "NOT_FOUND"}
    byid = dict((v["id"], v) for v in values)
    notes = []
    n_combo = 0
    for gid, g in groups.items():
        v = byid[gid]
        g["hits"] = _direct_hits(v, targets, p)
        ok = [h for h in g["hits"] if h["currency_check"] != "CONFLICT"]
        if ok or g["hits"]:
            g["status"] = "DIRECT"
            if not ok:
                g["conflict_only"] = True
    if p["bucket"]:
        todo = [gid for gid, g in groups.items() if g["status"] == "NOT_FOUND"]
        for gid in todo:
            if time.time() > deadline:
                groups[gid]["status"] = "SKIPPED"
                continue
            combos, strong = _combo_search(byid[gid], targets, p, deadline)
            if combos:
                groups[gid]["combos"] = combos
                groups[gid]["status"] = "COMBINATION" if strong else "POSSIBLE"
                n_combo += 1
        if any(g["status"] == "SKIPPED" for g in groups.values()):
            notes.append("The bucket (combination) search hit its time limit (%d s); the values marked 'Not searched' were skipped. "
                         "Raise the limit or lower the number of cells per bucket." % int(p["time_limit"]))
    palette_i = 0
    for g in groups.values():
        g["color"] = ""
        if g["status"] in ("DIRECT", "COMBINATION"):
            g["color"] = PALETTE[palette_i % len(PALETTE)]
            palette_i += 1
        elif g["status"] == "POSSIBLE":
            g["color"] = "FFF2CC"                                       # one pale colour for every weak (possible) combination
    stats = {"values": len(values), "distinct": len(groups), "direct": sum(1 for g in groups.values() if g["status"] == "DIRECT"),
             "combination": sum(1 for g in groups.values() if g["status"] == "COMBINATION"),
             "possible": sum(1 for g in groups.values() if g["status"] == "POSSIBLE"),
             "not_found": sum(1 for g in groups.values() if g["status"] == "NOT_FOUND"),
             "skipped": sum(1 for g in groups.values() if g["status"] == "SKIPPED"),
             "target_cells": sum(t["idx"].n_cells for t in targets), "seconds": round(time.time() - t0, 2)}
    for tg in targets:
        notes.extend(tg["idx"].warnings)
    for w in source["idx"].warnings:
        notes.append("Source: " + w)
    return {"params": p, "source": {"label": source["label"], "sheets": sorted(set(v["sheet"] for v in values))}, "values": values,
            "groups": groups, "targets": [t["label"] for t in targets], "stats": stats, "notes": notes}


def summary_rows(result, cap=5000):
    """One row per distinct source value for the web page."""
    byid = dict((v["id"], v) for v in result["values"])
    rows = []
    for gid, g in result["groups"].items():
        v = byid[gid]
        where = ""
        if g["hits"]:
            h = g["hits"][0]
            where = "%s | %s!%s" % (h["file"], h["sheet"], h["cell"])
        elif g["combos"]:
            c = g["combos"][0]
            where = "%s | %s: %s" % (c["file"], c["sheet"], " + ".join(x["cell"] for x in c["cells"]))
        rows.append({"id": gid, "sheet": v["sheet"], "cell": v["cell"], "row_label": v["row_label"], "column_header": v["column_header"],
                     "currency": v["currency"], "amount": v["amount"], "members": len(g["members"]), "status": g["status"],
                     "where": where, "n_hits": len(g["hits"]), "n_combos": len(g["combos"]), "color": g["color"],
                     "method": ("%s of %d cells" % (g["combos"][0]["tier"], g["combos"][0]["k"]) + ("" if g["combos"][0]["strong"] else
                                " - expected coincidences %s" % g["combos"][0]["chance"])) if g["combos"] else ("direct" if g["hits"] else ""),
                     "conflict_only": bool(g.get("conflict_only"))})
        if len(rows) >= cap:
            break
    return rows


# ------------------------------------------------------------------------------------------ annotated workbooks

def _note(lines, w=380):
    cm = Comment("\n".join(lines[:8]), "Depth engine")
    cm.width, cm.height = w, 40 + 18 * min(len(lines), 8)
    return cm


def _target_cells(result, label):
    """(sheet, cell) -> (colour, [comment lines]) for every cell found in this target file."""
    byid = dict((v["id"], v) for v in result["values"])
    out = OrderedDict()
    for gid, g in result["groups"].items():
        v = byid[gid]
        src = "%s = %s" % (v["row_label"] or v["cell"], format(v["amount"], ",.2f"))
        for h in g["hits"]:
            if h["file"] != label or h["currency_check"] == "CONFLICT":
                continue
            how = []
            if h["scale"] != 1.0:
                how.append("x%s" % format(h["scale"], "g"))
            if h["opposite_sign"]:
                how.append("opposite sign")
            out.setdefault((h["sheet"], h["cell"]), [g["color"], []])[1].append(
                "Source (%s!%s): %s%s" % (v["sheet"], v["cell"], src, (" - " + ", ".join(how)) if how else ""))
        for c in g["combos"][:1]:
            if c["file"] != label:
                continue
            for x in c["cells"]:
                out.setdefault((c["sheet"], x["cell"]), [g["color"], []])[1].append(
                    ("Part of a combination of %d cells (%s) that adds up to %s (source %s!%s)" if c["strong"] else
                     "POSSIBLE combination of %d cells (%s) adding up to %s (source %s!%s) - could be a coincidence") % (
                        c["k"], " + ".join(y["cell"] for y in c["cells"]), format(v["amount"], ",.2f"), v["sheet"], v["cell"]))
    return out


def write_target_workbook(path_in, path_out, label, result):
    wb = load_workbook(path_in, keep_vba=path_in.lower().endswith(".xlsm"))
    names = set(wb.sheetnames)
    bold = Font(bold=True)
    grey = _fill("D9D9D9")
    cells = _target_cells(result, label)
    for (sheet, cell), (colour, lines) in list(cells.items())[:5000]:
        if sheet in names:
            wb[sheet][cell].fill = _fill(colour)
            wb[sheet][cell].comment = _note(lines)
    values = result["values"]
    byid = dict((v["id"], v) for v in values)

    wv = wb.create_sheet(_unique_name(wb, "Source Values"))
    wv.append(["#", "Source sheet", "Cell", "Row label", "Column", "Currency", "Amount", "Status", "Shared with"])
    for c in wv[1]:
        c.font, c.fill = bold, grey
    for v in values:
        g = result["groups"][v["group"]]
        wv.append([v["id"] + 1, v["sheet"], v["cell"], v["row_label"], v["column_header"], v["currency"], v["amount"], STATUS_TEXT[g["status"]],
                   "" if v["group"] == v["id"] else "#%d" % (v["group"] + 1)])
        wv.cell(row=wv.max_row, column=7).number_format = "#,##0.00;[Red]-#,##0.00"
        if g["color"]:
            wv.cell(row=wv.max_row, column=1).fill = _fill(g["color"])
    for letter, w in zip("ABCDEFGHI", [6, 26, 8, 40, 36, 9, 20, 22, 10]):
        wv.column_dimensions[letter].width = w
    wv.freeze_panes = "A2"

    wr = wb.create_sheet(_unique_name(wb, "Matching Report"))
    head = ["Source #", "Source row", "Source column", "Source amount", "Found as", "Sheet", "Cell", "Value in sheet", "Unit", "Sign",
            "Row label", "Column", "Column currency", "Currency check", "Differs by", "Other hits"]
    wr.append(head)
    for c in wr[1]:
        c.font, c.fill = bold, grey
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for gid, g in result["groups"].items():
        v = byid[gid]
        for pos, h in enumerate(g["hits"]):
            if h["file"] != label:
                continue
            wr.append([gid + 1, v["row_label"], v["column_header"], v["amount"], "single cell", h["sheet"], h["cell"], h["value"],
                       "x%s" % format(h["scale"], "g"), "opposite" if h["opposite_sign"] else "same", h["row_label"], h["column_header"],
                       h["column_currency"], h["currency_check"], h["diff"], len(g["hits"]) - 1 if pos == 0 else ""])
            ri = wr.max_row
            if g["color"] and h["currency_check"] != "CONFLICT":
                wr.cell(row=ri, column=1).fill = _fill(g["color"])
            if h["sheet"] in names:
                cell = wr.cell(row=ri, column=7)
                cell.hyperlink = Hyperlink(ref=cell.coordinate, location="%s!%s" % (_sheet_ref(h["sheet"]), h["cell"]), display=h["cell"])
                cell.font = Font(color="0563C1", underline="single")
        for c in g["combos"][:1]:
            if c["file"] != label:
                continue
            for x in c["cells"]:
                wr.append([gid + 1, v["row_label"], v["column_header"], v["amount"],
                           "%scombination of %d cells (%s, %s %s%s)" % ("" if c["strong"] else "POSSIBLE ", c["k"], c["tier"], c["pool"], c["pool_key"],
                                                                          "" if c["strong"] else ", expected coincidences %s" % c["chance"]),
                           c["sheet"], x["cell"], x["value"], "x%s" % format(x["scale"], "g"), "opposite" if c["opposite_sign"] else "same",
                           x["row_label"], x["column_header"], x["column_currency"], "", c["diff"], ""])
                if g["color"]:
                    wr.cell(row=wr.max_row, column=1).fill = _fill(g["color"])
    for col in (4, 8, 15):
        for row in wr.iter_rows(min_row=2, min_col=col, max_col=col):
            row[0].number_format = "#,##0.00;[Red]-#,##0.00"
    for letter, w in zip("ABCDEFGHIJKLMNOP", [9, 36, 30, 18, 26, 26, 8, 18, 9, 9, 36, 30, 12, 12, 12, 10]):
        wr.column_dimensions[letter].width = w
    wr.freeze_panes = "A2"
    wr.auto_filter.ref = "A1:P%d" % max(wr.max_row, 2)

    wu = wb.create_sheet(_unique_name(wb, "Unmatched Source Values"))
    wu.append(["#", "Source sheet", "Cell", "Row label", "Column", "Currency", "Amount"])
    for c in wu[1]:
        c.font, c.fill = bold, grey
    for gid, g in result["groups"].items():
        if g["status"] in ("NOT_FOUND", "SKIPPED", "POSSIBLE"):
            v = byid[gid]
            wu.append([gid + 1, v["sheet"], v["cell"], v["row_label"], v["column_header"], v["currency"], v["amount"]])
            wu.cell(row=wu.max_row, column=7).number_format = "#,##0.00;[Red]-#,##0.00"
    for letter, w in zip("ABCDEFG", [6, 26, 8, 40, 36, 9, 20]):
        wu.column_dimensions[letter].width = w

    ws = wb.create_sheet(_unique_name(wb, "Depth Engine Summary"))
    p, st = result["params"], result["stats"]
    for a, b in [("Target file", label), ("Source file", result["source"]["label"]), ("Source sheets", ", ".join(result["source"]["sheets"])),
                 ("Generated", datetime.datetime.now().strftime("%Y-%m-%d %H:%M")), ("", ""),
                 ("Source numbers", st["values"]), ("Distinct values searched", st["distinct"]), ("Found in a single cell", st["direct"]),
                 ("Found as a combination of cells", st["combination"]), ("Possible combinations (could be a coincidence)", st["possible"]),
                 ("Not found", st["not_found"]), ("", ""),
                 ("Tolerance (absolute)", p["tol_abs"]), ("Thousands / millions tried", p["allow_scale"]), ("Opposite sign accepted", p["allow_sign"]),
                 ("Bucket (combination) search", p["bucket"]), ("Bucket mode", p["bucket_mode"]), ("Cells per bucket", p["bucket_cells"]),
                 ("Largest combination", p["bucket_max_k"]), ("Coincidences tolerated for a strong match", p["max_chance"])]:
        ws.append([a, b])
    for i in range(1, ws.max_row + 1):
        ws.cell(row=i, column=1).font = bold
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 50
    wb.save(path_out)
    return path_out


def write_source_workbook(path_in, path_out, label, result):
    """The source with every searched number coloured: the colour it shares with the target cells, or red when nothing was found."""
    wb = load_workbook(path_in, keep_vba=path_in.lower().endswith(".xlsm"))
    names = set(wb.sheetnames)
    byid = dict((v["id"], v) for v in result["values"])
    for v in result["values"]:
        g = result["groups"][v["group"]]
        if v["sheet"] not in names:
            continue
        cell = wb[v["sheet"]][v["cell"]]
        lines = []
        if g["status"] in ("DIRECT", "COMBINATION", "POSSIBLE"):
            cell.fill = _fill(g["color"])
            for h in g["hits"][:3]:
                if h["currency_check"] != "CONFLICT":
                    lines.append("Found in %s | %s!%s (%s)" % (h["file"], h["sheet"], h["cell"], h["row_label"]))
            for c in g["combos"][:1]:
                lines.append("%s %d cells in %s | %s: %s" % ("Found as" if c["strong"] else "POSSIBLY (could be a coincidence):", c["k"], c["file"], c["sheet"],
                                                              " + ".join(x["cell"] for x in c["cells"])))
        else:
            cell.fill = _fill("F4B6B6")
            lines.append("Not found in the target files" if g["status"] == "NOT_FOUND" else "Not searched (time limit)")
        cell.comment = _note(lines or ["Searched"])
    wb.save(path_out)
    return path_out


# ------------------------------------------------------------------------------------------ mapping

def _key_ccy(text, ccy):
    """The currency key of a source column (IQD, USD ... FRX, TOTAL) from its detected currency or its header text."""
    t = str(text or "").strip().lower()
    if t.startswith("frx"):
        return FRX_CODE
    if "grand total" in t or t == "total":
        return "TOTAL"
    return ccy or ""


def build_mapping(result, as_tb=False, only_unique=False):
    """The portable depth mapping: which source number sits where in the targets, by LABELS (row label, column header, currency),
    never by cell address, so it can be re-applied to the next period's files."""
    byid = dict((v["id"], v) for v in result["values"])
    entries = []
    for gid, g in result["groups"].items():
        v = byid[gid]
        if g["status"] == "DIRECT":
            ok = [h for h in g["hits"] if h["currency_check"] != "CONFLICT"]
            if not ok or (only_unique and len(ok) > 1):
                continue
            for member in g["members"][:1]:
                h = ok[0]
                entries.append({"source": {"sheet": v["sheet"], "row_label": v["row_label"], "column_header": v["column_header"],
                                           "currency": _key_ccy(v["column_header"], v["currency"])},
                                "amount": v["amount"], "method": "DIRECT", "alternatives": len(ok) - 1,
                                "targets": [{"file": h["file"], "sheet": h["sheet"], "row_label": h["row_label"],
                                             "column_header": h["column_header"], "currency": h["column_currency"],
                                             "scale": h["scale"], "sign": -1 if h["opposite_sign"] else 1}]})
        elif g["status"] == "COMBINATION" and g["combos"][0]["strong"]:
            c = g["combos"][0]
            entries.append({"source": {"sheet": v["sheet"], "row_label": v["row_label"], "column_header": v["column_header"],
                                       "currency": _key_ccy(v["column_header"], v["currency"])},
                            "amount": v["amount"], "method": "COMBINATION", "alternatives": len(g["combos"]) - 1,
                            "targets": [{"file": c["file"], "sheet": c["sheet"], "row_label": x["row_label"], "column_header": x["column_header"],
                                         "currency": x["column_currency"], "scale": x["scale"], "sign": -1 if c["opposite_sign"] else 1}
                                        for x in c["cells"]]})
    return {"format": "bahrain-iraq-depth-mapping", "version": 1, "created": datetime.datetime.now().isoformat(timespec="seconds"),
            "source": result["source"], "source_rows_are_bs_groups": bool(as_tb), "entries": entries}


def _sub_item(t, fallback_ccy):
    item = {"match_text": t["row_label"], "sheet": t["sheet"], "is_total": False, "sign": t.get("sign", 1)}
    if t.get("currency"):
        item["currency"] = t["currency"]
    else:
        item["currency"] = fallback_ccy or "TOTAL"
        if t.get("column_header"):
            item["column_context"] = t["column_header"]               # the column cannot be told by currency: one cell by its heading
    return item


def templates_from_mapping(doc, include_combinations=False):
    """Default-mapping rules from a depth mapping whose source rows are BS-mapping groups (a Trial Balance pivot): the group and
    currency on the Trial Balance side, the matched row(s) and column currency on the submission side."""
    out, seen = [], set()
    for e in doc.get("entries", []):
        if e.get("method") == "COMBINATION" and not include_combinations:
            continue
        s = e["source"]
        group = (s.get("row_label") or "").strip()
        ccy = s.get("currency") or "TOTAL"
        if not group:
            continue
        subs = [_sub_item(t, ccy) for t in e["targets"]]
        key = (group.lower(), ccy, tuple((x["sheet"], x["match_text"].lower(), x.get("currency")) for x in subs))
        if key in seen:
            continue
        seen.add(key)
        out.append({"label": "Depth search: %s | %s -> %s" % (group, ccy, " + ".join(t["row_label"] for t in e["targets"][:3])),
                    "currency": ccy, "rule_type": "GROUP_RULE", "source": "DEPTH_SEARCH", "tb_side": [
                        {"level": "group", "bs_mapping": group, "currency": ccy, "sign": 1}],
                    "submission_side": subs})
    return out


def templates_from_tb_depth(res):
    """Default-mapping rules from the Trial Balance depth search (Depth Search page): every BS-mapping value that was found in a
    submission becomes a group rule - TB group + currency -> the sheet, row and column currency where it was found."""
    T = dict((t["id"], t) for t in res["targets"])
    out, seen = [], set()
    for f in res["files"]:
        best = {}
        for h in f["hits"]:
            t = T.get(h["target_id"])
            if not t or h["currency_check"] == "CONFLICT" or t.get("external") or t["level"] != "BS Mapping":
                continue
            best.setdefault(h["target_id"], h)
        for tid, h in best.items():
            t = T[tid]
            ccy = t["currency"]
            item = _sub_item({"row_label": h["row_label"], "sheet": h["sheet"], "currency": h.get("column_currency") if h.get("column_currency") not in ("", None) else "",
                              "column_header": h["column_header"], "sign": -1 if h["opposite_sign"] else 1}, ccy)
            key = (t["label"].lower(), ccy, item["sheet"], item["match_text"].lower(), item.get("currency"))
            if key in seen:
                continue
            seen.add(key)
            out.append({"label": "Depth search: %s | %s -> %s" % (t["label"], ccy, h["row_label"]), "currency": ccy, "rule_type": "GROUP_RULE",
                        "source": "DEPTH_SEARCH", "tb_side": [{"level": "group", "bs_mapping": t["label"], "currency": ccy, "sign": 1}],
                        "submission_side": [item]})
    return out


def merge_templates(existing, new):
    """The existing default mapping plus the new rules that are not already in it (same TB group items and submission row)."""
    from .default_report import _rule_key
    have = set(k for k in (_rule_key(t) for t in existing) if k)
    out = list(existing)
    added = 0
    for t in new:
        k = _rule_key(t)
        if k and k in have:
            continue
        out.append(t)
        added += 1
        if k:
            have.add(k)
    return out, added


def merged_document(templates, note=""):
    """default_mapping.json with its `matches` replaced by the merged list - every other key (ob_row_groups ...) is kept, so the
    downloaded file can replace the one in the library as it is."""
    import json
    import os
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "default_mapping.json")
    try:
        with open(path, encoding="utf-8") as fh:
            doc = json.load(fh)
    except (OSError, ValueError):
        doc = {"format": "bahrain-iraq-recon-mapping", "version": 2}
    doc["matches"] = templates
    doc["generated_from"] = (str(doc.get("generated_from") or "") + " + " + note).strip(" +")
    return doc


def mapping_document(templates, note=""):
    return {"format": "bahrain-iraq-recon-mapping", "version": 2, "generated_from": note or "depth search", "matches": templates}
