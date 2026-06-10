
import io, json, math, re
import numpy as np
import pandas as pd
import streamlit as st
import cvxpy as cp

IDC = None  # set after model load

# Streamlit >=1.46 replaced use_container_width with width="stretch"; support both.
def _v(ver):
    return tuple(int(x) for x in re.findall(r"\d+", ver)[:2])
FULLW = {"width": "stretch"} if _v(st.__version__) >= (1, 46) else {"use_container_width": True}


# ============================================================================
# PARSING — native workbook layout (titled blocks, left/right pairs)
# ============================================================================
BLOCK_PATTERNS = [
    ("adj", r"ADJUSTED\s+UNIVERSE"),
    ("uni", r"^UNIVERSE"),
    ("fF",  r"FRESH.*SAMPLE\s+FRAME"), ("fP", r"PANEL.*SAMPLE\s+FRAME"),
    ("uF",  r"FRESH.*USED\s+CONTACTS"), ("uP", r"PANEL.*USED\s+CONTACTS"),
    ("cF",  r"FRESH.*COMPLETED"), ("cP", r"PANEL.*COMPLETED"),
    ("pF",  r"FRESH.*PREVIOUS"), ("pP", r"PANEL.*PREVIOUS"),
]
PARAM_PATTERNS = {
    "total":   r"^sample\s+size",
    "max_cell": r"^maximum\s+cell\s+size",
    "min_cell": r"^minimum\s+cell\s+size",
    "max_bw":  r"max(imum)?\s+base\s*-?\s*weight|desired\s+maximum\s+base\s*weight",
    "conv":    r"conversion\s+rate",
    "same":    r"same\s+source",
    "dedup":   r"de[\s\-]?dup",
}

def _cellstr(v):
    if v is None: return ""
    if isinstance(v, float) and math.isnan(v): return ""
    return str(v).strip()

def _read_block(grid, r0, c0):
    """Read one titled block: title at (r0,c0); header at r0+1 (c0 blank, c0+1='size', sectors...);
       data rows until region col empty. Returns (sectors, rows[(region,size,values)])."""
    hdr = grid[r0 + 1]
    sectors = []
    c = c0 + 2
    while c < len(hdr) and _cellstr(hdr[c]) != "":
        sectors.append(_cellstr(hdr[c])); c += 1
    if not sectors:
        return None
    rows, r = [], r0 + 2
    while r < len(grid):
        region = _cellstr(grid[r][c0]) if c0 < len(grid[r]) else ""
        size = _cellstr(grid[r][c0 + 1]) if c0 + 1 < len(grid[r]) else ""
        if region == "" or size == "":
            break
        vals = []
        for k in range(len(sectors)):
            v = grid[r][c0 + 2 + k] if c0 + 2 + k < len(grid[r]) else None
            try:
                fv = float(v) if v not in (None, "") else 0.0
                vals.append(0.0 if math.isnan(fv) else fv)
            except (TypeError, ValueError):
                vals.append(0.0)
        rows.append((region, size, vals)); r += 1
    return (sectors, rows) if rows else None

def parse_native(xls: dict):
    """Scan sheets for titled blocks + the parameter block. Sheets are ranked by how many
       distinct patterns they contain, so template/README sheets can't hijack the parse."""
    sheet_blocks, sheet_params = {}, {}
    for name, df in xls.items():
        grid = df.where(pd.notna(df), None).values.tolist()
        blocks, params = {}, {}
        for r, row in enumerate(grid):
            for c, v in enumerate(row):
                s = _cellstr(v)
                if not s:
                    continue
                up, low = s.upper(), s.lower()
                for key, pat in BLOCK_PATTERNS:
                    if key not in blocks and re.search(pat, up):
                        blk = _read_block(grid, r, c)
                        if blk:
                            blocks[key] = blk
                        break
                for pk, pat in PARAM_PATTERNS.items():
                    if pk not in params and re.search(pat, low):
                        val = row[c + 1] if c + 1 < len(row) else None
                        if val is not None and _cellstr(val) != "":
                            params[pk] = val
        if blocks:
            sheet_blocks[name] = blocks
        if params:
            sheet_params[name] = params
    if not sheet_blocks:
        return None
    # rank sheets: most distinct blocks; tie-break by name (Inputs > other > README/template-ish)
    # and by position (later sheets beat leading template copies like 'All Turkey')
    sheet_pos = {n: i for i, n in enumerate(xls)}
    def _prio(n):
        ln = n.lower()
        if "input" in ln: return 3
        if re.search(r"read\s*me|template|^all\s", ln): return 0
        return 1
    order = sorted(sheet_blocks, key=lambda n: (-len(sheet_blocks[n]), -_prio(n), -sheet_pos[n]))
    blocks = {}
    for name in order:
        for key, blk in sheet_blocks[name].items():
            blocks.setdefault(key, blk)
    if "fF" not in blocks or "fP" not in blocks or not ({"uni", "adj"} & set(blocks)):
        return None
    # rank param sheets: must have sample size; prefer most parameters (the Optimization sheet wins)
    params = {}
    porder = sorted(sheet_params, key=lambda n: (-int("total" in sheet_params[n]), -len(sheet_params[n])))
    for name in porder:
        for k, v in sheet_params[name].items():
            params.setdefault(k, v)
    sectors, base_rows = blocks["fF"]
    recs = {}
    order = []
    for key, (secs, rows) in blocks.items():
        for region, size, vals in rows:
            for sec, v in zip(secs, vals):
                k = (region, size, sec)
                if k not in recs:
                    recs[k] = {}; order.append(k)
                recs[k][key] = v
    # order by frame_fresh block, then leftovers
    ordered = [ (rg, sz, sc) for rg, sz, vals in base_rows for sc in sectors ]
    ordered += [k for k in order if k not in set(ordered)]
    rows = []
    for (rg, sz, sc) in ordered:
        d = recs.get((rg, sz, sc), {})
        rows.append({"Region": rg, "Size": sz, "Sector": sc, **{f: d.get(f, np.nan) for f, _ in BLOCK_PATTERNS}})
    df = pd.DataFrame(rows)
    if "uni" not in blocks:
        df["uni"] = df["adj"]
    df[["uni", "fF", "fP"]] = df[["uni", "fF", "fP"]].fillna(0)
    P = {}
    for k, v in params.items():
        if k in ("same", "dedup"):
            P[k] = str(v).strip().lower().startswith("y")
        else:
            try: P[k] = float(v)
            except (TypeError, ValueError): pass
    return {"id_cols": ["Region", "Size"], "df": df, "found": {"native": True},
            "file_params": P,
            "has": {"adj_sheet": "adj" in blocks, "used": "uF" in blocks or "uP" in blocks,
                    "completed": "cF" in blocks or "cP" in blocks,
                    "prev": "pF" in blocks or "pP" in blocks}}

# ----------------------------------------------------------------------------
# Clean wide-sheet fallback (one table per sheet)
# ----------------------------------------------------------------------------
SHEET_ALIASES = {
    "adj": ["adjusted_universe", "adjusted universe"], "uni": ["universe", "universe_external"],
    "fF": ["frame_fresh", "fresh_frame", "fresh"], "fP": ["frame_panel", "panel_frame", "panel"],
    "uF": ["used_fresh", "fresh_used"], "uP": ["used_panel", "panel_used"],
    "cF": ["completed_fresh", "achieved_fresh"], "cP": ["completed_panel", "achieved_panel"],
    "pF": ["previous_design_fresh", "prev_fresh"], "pP": ["previous_design_panel", "prev_panel"],
}
def _norm(s): return str(s).lower().replace(" ", "_").replace("-", "_").replace(":", "")

def parse_wide(xls: dict):
    frames, id_cols = {}, None
    for fld, aliases in SHEET_ALIASES.items():
        hit = next((n for n in xls if _norm(n) in map(_norm, aliases)), None)
        if hit is None:
            continue
        df = xls[hit].dropna(how="all").dropna(axis=1, how="all")
        n_id = 0
        for col in df.columns:
            vals = df[col].dropna()
            numish = vals.apply(lambda v: isinstance(v, (int, float))).mean() if len(vals) else 0
            if numish < 0.5: n_id += 1
            else: break
        if n_id == 0 or n_id >= len(df.columns):
            continue
        ids = list(df.columns[:n_id]); id_cols = id_cols or ids
        long = df.melt(id_vars=ids, var_name="Sector", value_name=fld)
        long[fld] = pd.to_numeric(long[fld], errors="coerce")
        for c in ids: long[c] = long[c].astype(str).str.strip()
        frames[fld] = long
    if "fF" not in frames or "fP" not in frames or not ({"uni", "adj"} & set(frames)):
        return None
    keys = id_cols + ["Sector"]
    out = frames["fF"][keys].drop_duplicates()
    for fld, f in frames.items():
        out = out.merge(f[keys + [fld]], on=keys, how="outer")
    for fld in SHEET_ALIASES:
        if fld not in out.columns: out[fld] = np.nan
    if "uni" not in frames: out["uni"] = out["adj"]
    out[["uni", "fF", "fP"]] = out[["uni", "fF", "fP"]].fillna(0)
    return {"id_cols": id_cols, "df": out, "found": {"wide": True}, "file_params": {},
            "has": {"adj_sheet": "adj" in frames, "used": "uF" in frames or "uP" in frames,
                    "completed": "cF" in frames or "cP" in frames, "prev": "pF" in frames or "pP" in frames}}

def parse_workbook(file):
    xls = pd.read_excel(file, sheet_name=None, header=None)
    model = parse_native(xls)
    if model is None:
        xls2 = pd.read_excel(file, sheet_name=None)
        model = parse_wide(xls2)
    return model or {"error": "Could not find the data blocks. Expected either the production layout "
                              "(Inputs sheet with titled blocks) or one wide sheet per table."}

def adjusted_universe(df, same_source, deduped):
    u, f, p = df["uni"], df["fF"].fillna(0), df["fP"].fillna(0)
    if same_source:
        return (f + p) if deduped else pd.concat([u, f, p], axis=1).max(axis=1)
    return pd.concat([u, f + p], axis=1).max(axis=1) if deduped else pd.concat([u, f, p], axis=1).max(axis=1)

# ============================================================================
# SOLVER — cvxpy MIP: SCIP -> ECOS_BB -> relax+round (original app's stack)
# ============================================================================
def n_infinity(z, moe, p): return z * z * p * (1 - p) / (moe * moe)
def fpc_min(N, ninf): return 0.0 if N <= 0 else ninf / (1 + ninf / N)

def _repair_sum(xr, lb, ub, total):
    delta = int(total - xr.sum())
    while delta != 0:
        room = (ub - xr) if delta > 0 else (xr - lb)
        i = int(np.argmax(room))
        if room[i] <= 0: break
        step = min(abs(delta), int(room[i]))
        xr[i] += step if delta > 0 else -step
        delta += -step if delta > 0 else step
    return xr

def solve_allocation(t, lb, ub, total, groups, solver_choice="SCIP", time_limit=20):
    n = len(t)
    x = cp.Variable(n, integer=True)
    cons = [x >= lb, x <= ub, cp.sum(x) == total] + \
           [cp.sum(x[g["members"]]) >= g["min"] for g in groups if g["min"] > 0]
    prob = cp.Problem(cp.Minimize(cp.sum_squares(x - t)), cons)
    order = [s for s in ([solver_choice] + ["SCIP", "ECOS_BB"]) if s in cp.installed_solvers()]
    seen, last_err = set(), None
    for s in order:
        if s in seen: continue
        seen.add(s)
        try:
            if s == "SCIP":
                prob.solve(solver=cp.SCIP, scip_params={"limits/time": float(time_limit)}, verbose=False)
            else:
                prob.solve(solver=cp.ECOS_BB, mi_max_iters=100000, verbose=False)
            if prob.status not in ("infeasible", "unbounded") and x.value is not None:
                xr = _repair_sum(np.clip(np.rint(np.asarray(x.value).ravel()), lb, ub).astype(int), lb, ub, total)
                if xr.sum() == total:
                    return xr, s, []
        except Exception as e:
            last_err = e
    # relax + round
    xr_v = cp.Variable(n)
    pr = cp.Problem(cp.Minimize(cp.sum_squares(xr_v - t)),
                    [xr_v >= lb, xr_v <= ub, cp.sum(xr_v) == total] +
                    [cp.sum(xr_v[g["members"]]) >= g["min"] for g in groups if g["min"] > 0])
    try:
        try: pr.solve(solver=cp.OSQP, verbose=False)
        except Exception: pr.solve(solver=cp.ECOS, verbose=False)
        if xr_v.value is None: raise RuntimeError(pr.status)
        xr = _repair_sum(np.clip(np.rint(xr_v.value), lb, ub).astype(int), lb, ub, total)
        for g in groups:
            short = int(g["min"] - xr[g["members"]].sum())
            if short > 0:
                slack = ub[g["members"]] - xr[g["members"]]
                for j in np.argsort(-slack):
                    if short == 0: break
                    add = min(short, int(slack[j]))
                    if add > 0: xr[g["members"][j]] += add; short -= add
                xr = _repair_sum(xr, lb, ub, total)
        if xr.sum() != total: raise RuntimeError("rounding missed the total")
        return xr, "RELAX+ROUND", ["MIP unavailable/failed — continuous relaxation + rounding used."]
    except Exception as e2:
        raise ValueError(f"No solver found a feasible solution (MIP: {last_err}; fallback: {e2})")

def slack_diagnostic(t, lb, ub, total, groups, df, idc):
    n = len(t)
    x = cp.Variable(n, nonneg=True); s_tot = cp.Variable(nonneg=True); cs = cp.Variable(n, nonneg=True)
    cons = [cp.sum(x) + s_tot == total, x <= ub, x + cs >= lb]
    ds = {}
    for g in groups:
        if g["min"] > 0:
            sv = cp.Variable(nonneg=True); ds[g["label"]] = sv
            cons.append(cp.sum(x[g["members"]]) + sv >= g["min"])
    try:
        cp.Problem(cp.Minimize(s_tot + cp.sum(cs) + sum(ds.values())), cons).solve(solver=cp.ECOS, verbose=False)
    except Exception as e:
        return pd.DataFrame([{"Constraint": "diagnostic error", "Slack needed": str(e)}])
    rows = []
    if s_tot.value and s_tot.value > 1e-6:
        rows.append({"Constraint": "Total sample", "Slack needed": round(float(s_tot.value), 1)})
    for lab, sv in ds.items():
        if sv.value and sv.value > 1e-6:
            rows.append({"Constraint": lab, "Slack needed": round(float(sv.value), 1)})
    csv_ = np.asarray(cs.value).ravel() if cs.value is not None else np.zeros(n)
    for i in np.where(csv_ > 1e-6)[0]:
        lab = " / ".join(str(df.iloc[i][c]) for c in idc + ["Sector"])
        rows.append({"Constraint": f"Cell min — {lab}", "Slack needed": round(float(csv_[i]), 1)})
    return pd.DataFrame(rows).sort_values("Slack needed", ascending=False) if rows else pd.DataFrame()

def split_add(add, capP, capF, policy, pshare, achP, achF, final):
    """Split ADDITIONAL interviews; per-cell panel cap = pshare * FINAL design (Target/2 rule)."""
    aP = np.zeros_like(add)
    for i, a in enumerate(add):
        room_p = max(0, int(math.floor(final[i] * pshare + 1e-9)) - achP[i])
        if policy == "Fresh first":
            p = max(0, a - min(a, capF[i]))
        elif policy == "Proportional to frames":
            tot = capP[i] + capF[i]
            p = int(round(a * capP[i] / tot)) if tot > 0 else 0
        else:
            p = min(a, capP[i])
        p = max(0, min(p, room_p, capP[i], a))
        if a - p > capF[i]:
            p = min(capP[i], a - capF[i])
        aP[i] = p
    return aP, add - aP

# ============================================================================
# PIPELINE — one solver: redesign-aware when fieldwork data exists
# ============================================================================
def run_solver(model, P, mins, ignore_fieldwork=False):
    df = model["df"].copy()
    idc = model["id_cols"]
    use_field = model["has"]["completed"] and not ignore_fieldwork
    df["pop"] = (df["uni"] if P["basis"].startswith("External")
                 else df["adj"].fillna(adjusted_universe(df, P["same"], P["dedup"])) if P["basis"].startswith("Uploaded")
                 else adjusted_universe(df, P["same"], P["dedup"]))
    cF = df["cF"].fillna(0) if use_field else pd.Series(0.0, index=df.index)
    cP = df["cP"].fillna(0) if use_field else pd.Series(0.0, index=df.index)
    uF = df["uF"].fillna(df["cF"].fillna(0)) if use_field else pd.Series(0.0, index=df.index)
    uP = df["uP"].fillna(df["cP"].fillna(0)) if use_field else pd.Series(0.0, index=df.index)
    df["cF_"], df["cP_"] = cF, cP
    df["ach"] = cF + cP
    df["capF"] = (df["fF"] - uF).clip(lower=0)
    df["capP"] = (df["fP"] - uP).clip(lower=0)
    df["avail"] = df["capF"] + df["capP"]
    tot_pop = df["pop"].sum() or 1.0
    df["t"] = df["pop"] * P["total"] / tot_pop

    # TOTAL bounds per cell (sheet semantics: MIN = max(completed, expected/min...), MAX accordingly)
    conv_room = np.ceil(df["avail"] * P["conv"])
    ub_tot = df["ach"] + np.minimum.reduce([df["avail"].values, conv_room.values,
                                            np.maximum(0, P["max_cell"] - df["ach"].values)])
    reach = df["avail"] + df["ach"]
    lb_base = np.where(reach > 0,
                       np.maximum(np.ceil(df["pop"] / P["max_bw"]) if P["max_bw"] > 0 else 0,
                                  np.where(np.minimum(reach, max(P["max_cell"], 1)) >= P["min_cell"], P["min_cell"], 0)),
                       0)
    lb_tot = np.maximum(df["ach"].values, np.minimum(lb_base, ub_tot))  # never below completed; clipped to reachable
    df["lb_tot"], df["ub_tot"] = lb_tot.astype(int), np.floor(ub_tot).astype(int)

    # groups on TOTALS
    groups = []
    for (dim, val), req in mins.items():
        if req <= 0: continue
        col = df["Sector"] if dim == "Sector" else df[dim]
        members = np.where(col.astype(str) == str(val))[0]
        if len(members):
            groups.append({"label": f"{dim} = {val}", "min": int(req), "members": members})

    # solve on additional interviews
    lb_add = (df["lb_tot"] - df["ach"]).clip(lower=0).astype(int).values
    ub_add = (df["ub_tot"] - df["ach"]).clip(lower=0).astype(int).values
    achieved = int(df["ach"].sum())
    R = max(0, int(P["total"]) - achieved)
    notes, conflicts = [], {"overall": [], "cells": [], "dims": []}
    g_add = []
    for g in groups:
        rem = max(0, g["min"] - int(df["ach"].iloc[g["members"]].sum()))
        cap = int(ub_add[g["members"]].sum())
        if rem > cap:
            if use_field:
                notes.append(f'{g["label"]}: remaining minimum {rem} exceeds capacity {cap} — clipped.')
                rem = cap
            else:
                conflicts["dims"].append((g["label"], g["min"], cap))
        if rem > 0:
            g_add.append({"label": g["label"], "min": rem, "members": g["members"]})
    bad = lb_add > ub_add
    if bad.any():
        if use_field:
            notes.append(f"{int(bad.sum())} cell minimum(s) unreachable from remaining contacts — relaxed.")
            lb_add = np.minimum(lb_add, ub_add)
        else:
            conflicts["cells"] = list(np.where(bad)[0])
    cap_all = int(ub_add.sum())
    if cap_all < R:
        if use_field:
            notes.append(f"Remaining frame supports only {cap_all} of {R} outstanding interviews — target clipped.")
            R = cap_all
        else:
            conflicts["overall"].append(("Frame + conversion capacity", int(P["total"]), cap_all + achieved))
    if int(lb_add.sum()) > R:
        if use_field:
            notes.append("Cell minimums jointly exceed the outstanding total — proportionally relaxed.")
            lb_add = np.floor(lb_add * R / max(1, lb_add.sum())).astype(int)
        else:
            conflicts["overall"].append(("Sum of cell lower bounds", int(lb_add.sum()) + achieved, int(P["total"])))
    if conflicts["overall"] or conflicts["cells"] or conflicts["dims"]:
        t_add = np.maximum(0, df["t"].values - df["ach"].values)
        return {"ok": False, "conflicts": conflicts, "df": df, "use_field": use_field,
                "slack": slack_diagnostic(t_add, lb_add, ub_add, R, g_add, df, idc)}

    t_add = np.maximum(0, df["t"].values - df["ach"].values)
    add, solver, sn = solve_allocation(t_add, lb_add, ub_add, R, g_add, P["solver"], P["time_limit"])
    notes += sn
    df["add"] = add
    df["x"] = (df["ach"] + df["add"]).astype(int)            # FULL SAMPLE DESIGN (final)
    aP, aF = split_add(add, df["capP"].values.astype(int), df["capF"].values.astype(int),
                       P["split"], P["pshare"], cP.values.astype(int), cF.values.astype(int), df["x"].values)
    df["panel"] = (cP + aP).astype(int)                       # PANEL SAMPLE DESIGN
    df["fresh"] = (cF + aF).astype(int)                       # FRESH SAMPLE DESIGN
    df["bw"] = np.where(df["x"] > 0, df["pop"] / df["x"].replace(0, np.nan), np.nan)
    # red-check tables: remaining contacts per source; red when < still needed from that source
    df["redF"] = df["capF"].astype(int)
    df["redP"] = df["capP"].astype(int)
    df["needF"] = (df["fresh"] - cF).clip(lower=0).astype(int)
    df["needP"] = (df["panel"] - cP).clip(lower=0).astype(int)
    if model["has"]["prev"] and use_field:
        df["overF"] = (cF - df["pF"].fillna(0)).clip(lower=0).astype(int)
        df["overP"] = (cP - df["pP"].fillna(0)).clip(lower=0).astype(int)
    return {"ok": True, "df": df, "groups": groups, "solver": solver, "notes": notes,
            "use_field": use_field, "achieved": achieved, "additional": int(df["add"].sum())}

# ============================================================================
# PRESENTATION — Excel Optimization-sheet style
# ============================================================================
CAP_CSS = {"green": ("#EAF3EB", "#2C5931", "#3E7C45"), "amber": ("#FCF2E2", "#7C4A00", "#C77800"),
           "violet": ("#EFEBF6", "#4A3F70", "#6B5B95"), "blue": ("#E7F0FA", "#0B4A82", "#0F6CBD")}

def cap_bar(title, theme, sub=""):
    bg, fg, bar = CAP_CSS[theme]
    st.markdown(
        f'<div style="background:{bg};border-left:5px solid {bar};border-radius:5px;'
        f'padding:5px 11px;margin:2px 0 6px;font-size:12px;font-weight:750;'
        f'letter-spacing:.05em;text-transform:uppercase;color:{fg}">{title}'
        f'{f"<span style=&quot;font-weight:500;text-transform:none;letter-spacing:0;opacity:.8&quot;> — {sub}</span>" if sub else ""}</div>',
        unsafe_allow_html=True)

def pivot(df, col, dec=0, totals_row=True):
    pv = df.pivot_table(index=IDC, columns="Sector", values=col, aggfunc="sum", sort=False)
    pv = pv[[s for s in df["Sector"].drop_duplicates() if s in pv.columns]]
    if totals_row:
        pv.loc[tuple(["TOTAL"] + [""] * (len(IDC) - 1)), :] = pv.sum(axis=0)
    return pv.round(dec)

def mini_caption(df, col):
    by_size = df.groupby(df[IDC[-1]], sort=False)[col].sum()
    by_reg = df.groupby(df[IDC[0]], sort=False)[col].sum()
    tot = df[col].sum()
    s1 = " · ".join(f"{k} **{v:,.0f}**" for k, v in by_size.items())
    s2 = " · ".join(f"{str(k)[:18]} **{v:,.0f}**" for k, v in by_reg.items())
    st.caption(f"{s1} — total **{tot:,.0f}**  \n{s2}")

def show_plain(df, col, dec=0, height=460):
    st.dataframe(pivot(df, col, dec).style.format(f"{{:,.{dec}f}}", na_rep="–"),
                 **FULLW, height=height)

def show_heat(df, col, dec=0, height=460):
    pv = pivot(df, col, dec, totals_row=False)
    st.dataframe(pv.style.format(f"{{:,.{dec}f}}", na_rep="–")
                 .background_gradient(cmap="RdYlGn_r", axis=None),
                 **FULLW, height=height)

def show_red(df, valcol, needcol, height=460):
    pv = pivot(df, valcol, totals_row=False)
    need = pivot(df, needcol, totals_row=False)
    def styler(data):
        out = pd.DataFrame("", index=data.index, columns=data.columns)
        mask = data < need.reindex_like(data)
        out[mask] = "background-color:#F4B4AE;color:#7A150F;font-weight:700"
        return out
    st.dataframe(pv.style.format("{:,.0f}", na_rep="–").apply(styler, axis=None),
                 **FULLW, height=height)

def show_flag_pos(df, col, height=460):
    pv = pivot(df, col)
    def f(v):
        if pd.isna(v): return ""
        return "background-color:#F4B4AE;color:#7A150F;font-weight:700" if v > 0 else "color:#B9C4CF"
    st.dataframe(pv.style.format("{:,.0f}", na_rep="–").map(f), **FULLW, height=height)

def to_excel(model, P, mins, res):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        pd.DataFrame([("Sample size", P["total"]), ("Min cell", P["min_cell"]), ("Max cell", P["max_cell"]),
                      ("Max base weight", P["max_bw"]), ("Conversion", P["conv"]), ("Max panel share", P["pshare"]),
                      ("Same source", P["same"]), ("De-duplicated", P["dedup"]), ("Basis", P["basis"]),
                      ("Split", P["split"]), ("Solver used", res.get("solver", ""))],
                     columns=["Parameter", "Value"]).to_excel(w, sheet_name="Parameters", index=False)
        pd.DataFrame([{"Dimension": d, "Value": v, "Min": m} for (d, v), m in mins.items()]
                     ).to_excel(w, sheet_name="Dimension_minimums", index=False)
        d = res["df"]
        sheets = [("Full_sample_design", "x", 0), ("Base_weights", "bw", 1),
                  ("Minimum_cell_size", "lb_tot", 0), ("Maximum_cell_size", "ub_tot", 0),
                  ("Fresh_sample_design", "fresh", 0), ("Panel_sample_design", "panel", 0),
                  ("Remaining_fresh_contacts", "redF", 0), ("Remaining_panel_contacts", "redP", 0)]
        if "overF" in d: sheets += [("Overshoot_fresh", "overF", 0), ("Overshoot_panel", "overP", 0)]
        for name, col, dec in sheets:
            pivot(d, col, dec).to_excel(w, sheet_name=name)
        d[model["id_cols"] + ["Sector", "pop", "fF", "fP", "cF_", "cP_", "lb_tot", "ub_tot",
                              "x", "fresh", "panel", "add", "bw"]].to_excel(w, sheet_name="Cells", index=False)
    out.seek(0)
    return out.getvalue()

DEMO_JSON = r"""{"idCols": ["Region", "Size"], "industries": ["Food Products", "Garment", "Other Manufacturing", "Retail", "Other Services"], "rows": [{"id": ["Western", "Small"], "ind": "Food Products", "adj": 725, "uni": 725, "fF": 395, "fP": 10, "uF": 10, "uP": 9, "cF": 2, "cP": 2, "pF": 2, "pP": 2}, {"id": ["Western", "Small"], "ind": "Garment", "adj": 1070, "uni": 1070, "fF": 591, "fP": 63, "uF": 2, "uP": 14, "cF": 2, "cP": 3, "pF": 2, "pP": 3}, {"id": ["Western", "Small"], "ind": "Other Manufacturing", "adj": 3256, "uni": 3256, "fF": 2249, "fP": 14, "uF": 75, "uP": 14, "cF": 21, "cP": 2, "pF": 20, "pP": 2}, {"id": ["Western", "Small"], "ind": "Retail", "adj": 3660, "uni": 3660, "fF": 1322, "fP": 10, "uF": 111, "uP": 10, "cF": 32, "cP": 1, "pF": 30, "pP": 1}, {"id": ["Western", "Small"], "ind": "Other Services", "adj": 9051, "uni": 9051, "fF": 2969, "fP": 10, "uF": 172, "uP": 10, "cF": 42, "cP": 1, "pF": 39, "pP": 1}, {"id": ["Western", "Medium"], "ind": "Food Products", "adj": 201, "uni": 201, "fF": 182, "fP": 9, "uF": 22, "uP": 9, "cF": 5, "cP": 5, "pF": 5, "pP": 5}, {"id": ["Western", "Medium"], "ind": "Garment", "adj": 367, "uni": 337, "fF": 342, "fP": 25, "uF": 22, "uP": 20, "cF": 6, "cP": 5, "pF": 6, "pP": 7}, {"id": ["Western", "Medium"], "ind": "Other Manufacturing", "adj": 1272, "uni": 1021, "fF": 1261, "fP": 11, "uF": 60, "uP": 11, "cF": 13, "cP": 3, "pF": 11, "pP": 3}, {"id": ["Western", "Medium"], "ind": "Retail", "adj": 540, "uni": 540, "fF": 297, "fP": 13, "uF": 10, "uP": 13, "cF": 4, "cP": 4, "pF": 4, "pP": 4}, {"id": ["Western", "Medium"], "ind": "Other Services", "adj": 1909, "uni": 1909, "fF": 1344, "fP": 10, "uF": 14, "uP": 10, "cF": 6, "cP": 3, "pF": 6, "pP": 3}, {"id": ["Western", "Large"], "ind": "Food Products", "adj": 104, "uni": 92, "fF": 98, "fP": 6, "uF": 30, "uP": 4, "cF": 8, "cP": 0, "pF": 13, "pP": 2}, {"id": ["Western", "Large"], "ind": "Garment", "adj": 356, "uni": 301, "fF": 346, "fP": 10, "uF": 35, "uP": 6, "cF": 7, "cP": 0, "pF": 15, "pP": 4}, {"id": ["Western", "Large"], "ind": "Other Manufacturing", "adj": 529, "uni": 387, "fF": 525, "fP": 4, "uF": 16, "uP": 3, "cF": 3, "cP": 2, "pF": 9, "pP": 3}, {"id": ["Western", "Large"], "ind": "Retail", "adj": 77, "uni": 50, "fF": 68, "fP": 9, "uF": 16, "uP": 9, "cF": 8, "cP": 3, "pF": 7, "pP": 3}, {"id": ["Western", "Large"], "ind": "Other Services", "adj": 455, "uni": 401, "fF": 442, "fP": 13, "uF": 7, "uP": 6, "cF": 1, "cP": 1, "pF": 2, "pP": 2}, {"id": ["Southern, Sabaragamuwa, Uva", "Small"], "ind": "Food Products", "adj": 497, "uni": 497, "fF": 402, "fP": 13, "uF": 8, "uP": 10, "cF": 2, "cP": 2, "pF": 2, "pP": 2}, {"id": ["Southern, Sabaragamuwa, Uva", "Small"], "ind": "Garment", "adj": 327, "uni": 327, "fF": 235, "fP": 0, "uF": 11, "uP": 0, "cF": 5, "cP": 0, "pF": 5, "pP": 0}, {"id": ["Southern, Sabaragamuwa, Uva", "Small"], "ind": "Other Manufacturing", "adj": 1379, "uni": 1379, "fF": 1094, "fP": 11, "uF": 7, "uP": 7, "cF": 2, "cP": 2, "pF": 2, "pP": 2}, {"id": ["Southern, Sabaragamuwa, Uva", "Small"], "ind": "Retail", "adj": 1599, "uni": 1599, "fF": 1012, "fP": 19, "uF": 6, "uP": 8, "cF": 3, "cP": 6, "pF": 3, "pP": 6}, {"id": ["Southern, Sabaragamuwa, Uva", "Small"], "ind": "Other Services", "adj": 3051, "uni": 3051, "fF": 1243, "fP": 14, "uF": 14, "uP": 11, "cF": 8, "cP": 8, "pF": 7, "pP": 8}, {"id": ["Southern, Sabaragamuwa, Uva", "Medium"], "ind": "Food Products", "adj": 339, "uni": 339, "fF": 283, "fP": 8, "uF": 39, "uP": 8, "cF": 10, "cP": 3, "pF": 9, "pP": 2}, {"id": ["Southern, Sabaragamuwa, Uva", "Medium"], "ind": "Garment", "adj": 99, "uni": 84, "fF": 99, "fP": 0, "uF": 29, "uP": 0, "cF": 10, "cP": 0, "pF": 10, "pP": 0}, {"id": ["Southern, Sabaragamuwa, Uva", "Medium"], "ind": "Other Manufacturing", "adj": 252, "uni": 202, "fF": 239, "fP": 13, "uF": 5, "uP": 6, "cF": 3, "cP": 3, "pF": 3, "pP": 3}, {"id": ["Southern, Sabaragamuwa, Uva", "Medium"], "ind": "Retail", "adj": 182, "uni": 141, "fF": 168, "fP": 14, "uF": 2, "uP": 11, "cF": 2, "cP": 2, "pF": 2, "pP": 2}, {"id": ["Southern, Sabaragamuwa, Uva", "Medium"], "ind": "Other Services", "adj": 356, "uni": 356, "fF": 333, "fP": 14, "uF": 10, "uP": 4, "cF": 1, "cP": 3, "pF": 1, "pP": 3}, {"id": ["Southern, Sabaragamuwa, Uva", "Large"], "ind": "Food Products", "adj": 159, "uni": 138, "fF": 151, "fP": 8, "uF": 101, "uP": 8, "cF": 14, "cP": 3, "pF": 15, "pP": 1}, {"id": ["Southern, Sabaragamuwa, Uva", "Large"], "ind": "Garment", "adj": 125, "uni": 109, "fF": 125, "fP": 0, "uF": 58, "uP": 0, "cF": 14, "cP": 0, "pF": 17, "pP": 0}, {"id": ["Southern, Sabaragamuwa, Uva", "Large"], "ind": "Other Manufacturing", "adj": 67, "uni": 55, "fF": 60, "fP": 7, "uF": 13, "uP": 7, "cF": 5, "cP": 1, "pF": 5, "pP": 1}, {"id": ["Southern, Sabaragamuwa, Uva", "Large"], "ind": "Retail", "adj": 26, "uni": 19, "fF": 26, "fP": 0, "uF": 23, "uP": 0, "cF": 7, "cP": 0, "pF": 6, "pP": 0}, {"id": ["Southern, Sabaragamuwa, Uva", "Large"], "ind": "Other Services", "adj": 44, "uni": 37, "fF": 41, "fP": 3, "uF": 10, "uP": 2, "cF": 1, "cP": 1, "pF": 2, "pP": 2}, {"id": ["Central, North-Central, North-Western", "Small"], "ind": "Food Products", "adj": 1155, "uni": 1155, "fF": 586, "fP": 20, "uF": 6, "uP": 6, "cF": 3, "cP": 3, "pF": 3, "pP": 3}, {"id": ["Central, North-Central, North-Western", "Small"], "ind": "Garment", "adj": 337, "uni": 337, "fF": 246, "fP": 11, "uF": 4, "uP": 2, "cF": 2, "cP": 0, "pF": 2, "pP": 2}, {"id": ["Central, North-Central, North-Western", "Small"], "ind": "Other Manufacturing", "adj": 2832, "uni": 2832, "fF": 1466, "fP": 21, "uF": 26, "uP": 18, "cF": 7, "cP": 9, "pF": 8, "pP": 9}, {"id": ["Central, North-Central, North-Western", "Small"], "ind": "Retail", "adj": 2184, "uni": 2184, "fF": 868, "fP": 26, "uF": 11, "uP": 14, "cF": 5, "cP": 6, "pF": 6, "pP": 7}, {"id": ["Central, North-Central, North-Western", "Small"], "ind": "Other Services", "adj": 4225, "uni": 4225, "fF": 1454, "fP": 18, "uF": 27, "uP": 15, "cF": 12, "cP": 7, "pF": 12, "pP": 10}, {"id": ["Central, North-Central, North-Western", "Medium"], "ind": "Food Products", "adj": 327, "uni": 327, "fF": 276, "fP": 21, "uF": 8, "uP": 13, "cF": 5, "cP": 4, "pF": 5, "pP": 6}, {"id": ["Central, North-Central, North-Western", "Medium"], "ind": "Garment", "adj": 157, "uni": 157, "fF": 147, "fP": 8, "uF": 17, "uP": 2, "cF": 5, "cP": 2, "pF": 5, "pP": 5}, {"id": ["Central, North-Central, North-Western", "Medium"], "ind": "Other Manufacturing", "adj": 410, "uni": 380, "fF": 391, "fP": 19, "uF": 3, "uP": 4, "cF": 2, "cP": 3, "pF": 2, "pP": 2}, {"id": ["Central, North-Central, North-Western", "Medium"], "ind": "Retail", "adj": 222, "uni": 222, "fF": 166, "fP": 25, "uF": 6, "uP": 7, "cF": 1, "cP": 2, "pF": 2, "pP": 3}, {"id": ["Central, North-Central, North-Western", "Medium"], "ind": "Other Services", "adj": 444, "uni": 444, "fF": 355, "fP": 16, "uF": 4, "uP": 4, "cF": 2, "cP": 4, "pF": 2, "pP": 4}, {"id": ["Central, North-Central, North-Western", "Large"], "ind": "Food Products", "adj": 125, "uni": 103, "fF": 114, "fP": 11, "uF": 27, "uP": 10, "cF": 3, "cP": 2, "pF": 11, "pP": 4}, {"id": ["Central, North-Central, North-Western", "Large"], "ind": "Garment", "adj": 143, "uni": 134, "fF": 140, "fP": 3, "uF": 23, "uP": 1, "cF": 11, "cP": 1, "pF": 14, "pP": 3}, {"id": ["Central, North-Central, North-Western", "Large"], "ind": "Other Manufacturing", "adj": 113, "uni": 90, "fF": 105, "fP": 8, "uF": 6, "uP": 6, "cF": 1, "cP": 2, "pF": 3, "pP": 4}, {"id": ["Central, North-Central, North-Western", "Large"], "ind": "Retail", "adj": 21, "uni": 13, "fF": 21, "fP": 0, "uF": 18, "uP": 0, "cF": 5, "cP": 0, "pF": 5, "pP": 0}, {"id": ["Central, North-Central, North-Western", "Large"], "ind": "Other Services", "adj": 92, "uni": 65, "fF": 84, "fP": 8, "uF": 1, "uP": 7, "cF": 1, "cP": 0, "pF": 2, "pP": 2}, {"id": ["Northern, Eastern", "Small"], "ind": "Food Products", "adj": 483, "uni": 483, "fF": 209, "fP": 7, "uF": 12, "uP": 7, "cF": 7, "cP": 4, "pF": 6, "pP": 4}, {"id": ["Northern, Eastern", "Small"], "ind": "Garment", "adj": 114, "uni": 114, "fF": 64, "fP": 0, "uF": 15, "uP": 0, "cF": 12, "cP": 0, "pF": 12, "pP": 0}, {"id": ["Northern, Eastern", "Small"], "ind": "Other Manufacturing", "adj": 621, "uni": 621, "fF": 448, "fP": 10, "uF": 8, "uP": 8, "cF": 4, "cP": 4, "pF": 4, "pP": 3}, {"id": ["Northern, Eastern", "Small"], "ind": "Retail", "adj": 829, "uni": 829, "fF": 290, "fP": 3, "uF": 15, "uP": 3, "cF": 10, "cP": 1, "pF": 10, "pP": 1}, {"id": ["Northern, Eastern", "Small"], "ind": "Other Services", "adj": 1499, "uni": 1499, "fF": 550, "fP": 9, "uF": 6, "uP": 8, "cF": 3, "cP": 5, "pF": 3, "pP": 5}, {"id": ["Northern, Eastern", "Medium"], "ind": "Food Products", "adj": 45, "uni": 36, "fF": 37, "fP": 8, "uF": 13, "uP": 8, "cF": 6, "cP": 6, "pF": 6, "pP": 7}, {"id": ["Northern, Eastern", "Medium"], "ind": "Garment", "adj": 4, "uni": 4, "fF": 2, "fP": 0, "uF": 2, "uP": 0, "cF": 2, "cP": 0, "pF": 2, "pP": 0}, {"id": ["Northern, Eastern", "Medium"], "ind": "Other Manufacturing", "adj": 64, "uni": 43, "fF": 60, "fP": 4, "uF": 56, "uP": 4, "cF": 17, "cP": 4, "pF": 8, "pP": 4}, {"id": ["Northern, Eastern", "Medium"], "ind": "Retail", "adj": 75, "uni": 75, "fF": 53, "fP": 8, "uF": 31, "uP": 7, "cF": 12, "cP": 3, "pF": 12, "pP": 4}, {"id": ["Northern, Eastern", "Medium"], "ind": "Other Services", "adj": 139, "uni": 139, "fF": 94, "fP": 5, "uF": 6, "uP": 5, "cF": 4, "cP": 3, "pF": 4, "pP": 3}, {"id": ["Northern, Eastern", "Large"], "ind": "Food Products", "adj": 2, "uni": 2, "fF": 2, "fP": 0, "uF": 2, "uP": 0, "cF": 1, "cP": 0, "pF": 2, "pP": 0}, {"id": ["Northern, Eastern", "Large"], "ind": "Garment", "adj": 22, "uni": 15, "fF": 22, "fP": 0, "uF": 22, "uP": 0, "cF": 5, "cP": 0, "pF": 5, "pP": 0}, {"id": ["Northern, Eastern", "Large"], "ind": "Other Manufacturing", "adj": 12, "uni": 8, "fF": 12, "fP": 0, "uF": 12, "uP": 0, "cF": 3, "cP": 0, "pF": 5, "pP": 0}, {"id": ["Northern, Eastern", "Large"], "ind": "Retail", "adj": 5, "uni": 3, "fF": 5, "fP": 0, "uF": 5, "uP": 0, "cF": 1, "cP": 0, "pF": 1, "pP": 0}, {"id": ["Northern, Eastern", "Large"], "ind": "Other Services", "adj": 20, "uni": 20, "fF": 12, "fP": 0, "uF": 12, "uP": 0, "cF": 3, "cP": 0, "pF": 6, "pP": 0}]}"""
def load_demo():
    d = json.loads(DEMO_JSON)
    rows = [dict(zip(d["idCols"], r["id"]), Sector=r["ind"], adj=r["adj"], uni=r["uni"],
                 fF=r["fF"], fP=r["fP"], uF=r["uF"], uP=r["uP"], cF=r["cF"], cP=r["cP"], pF=r["pF"], pP=r["pP"])
            for r in d["rows"]]
    return {"id_cols": d["idCols"], "df": pd.DataFrame(rows), "found": {"demo": True},
            "file_params": {"total": 600, "max_cell": 40, "min_cell": 2, "max_bw": 1000, "conv": 0.31,
                            "same": False, "dedup": True},
            "has": {"adj_sheet": True, "used": True, "completed": True, "prev": True}}

# ============================================================================
# UI
# ============================================================================
st.set_page_config(page_title="WBES Survey Design Studio", layout="wide", initial_sidebar_state="expanded")
st.markdown("""<style>
.block-container{padding-top:1rem;max-width:1750px}
h1{font-size:1.45rem!important}
[data-testid="stMetricValue"]{font-size:1.4rem}
div[data-testid="stDataFrame"] *{font-variant-numeric:tabular-nums}
.stTabs [data-baseweb="tab"]{font-weight:600}
</style>""", unsafe_allow_html=True)

DEFAULTS = dict(total=600, min_cell=2, max_cell=40, max_bw=1000, conv=0.31, pshare=50)
for k, v in DEFAULTS.items():
    st.session_state.setdefault(k, v)
st.session_state.setdefault("same_lbl", "No"); st.session_state.setdefault("dedup_lbl", "Yes")
st.session_state.setdefault("model", None); st.session_state.setdefault("res", None)
st.session_state.setdefault("mins", {})

def apply_file_params(fp):
    cast = {"total": int, "min_cell": int, "max_cell": int, "max_bw": float, "conv": float}
    for k, fn in cast.items():
        if k in fp: st.session_state[k] = fn(fp[k])
    if "same" in fp: st.session_state["same_lbl"] = "Yes" if fp["same"] else "No"
    if "dedup" in fp: st.session_state["dedup_lbl"] = "Yes" if fp["dedup"] else "No"

with st.sidebar:
    st.markdown("### Data")
    up = st.file_uploader("Excel workbook", type=["xlsx", "xls"], label_visibility="collapsed")
    if st.button("Load demo (Sri Lanka)", **FULLW):
        st.session_state.model = load_demo(); st.session_state.res = None
        apply_file_params(st.session_state.model["file_params"]); st.rerun()
    if up is not None and st.session_state.get("_upname") != up.name:
        m = parse_workbook(up)
        if "error" in m:
            st.error(m["error"])
        else:
            st.session_state.model = m; st.session_state._upname = up.name; st.session_state.res = None
            apply_file_params(m.get("file_params", {})); st.rerun()
    model = st.session_state.model
    if model:
        kind = "production layout" if "native" in model["found"] else ("demo" if "demo" in model["found"] else "wide sheets")
        st.success(f"{len(model['df'])} cells · {model['df']['Sector'].nunique()} sectors · {kind}")
        if model.get("file_params"):
            st.caption("Parameters auto-read from the workbook ✓")

    st.markdown("### Parameters")
    st.number_input("sample size", 1, 100000, key="total")
    c1, c2 = st.columns(2)
    c1.number_input("Minimum cell size", 0, 1000, key="min_cell")
    c2.number_input("Maximum cell size", 1, 10000, key="max_cell")
    st.number_input("Desired maximum base weight", 1, 100000, key="max_bw")
    st.number_input("Interview conversion rate", 0.01, 1.0, key="conv", step=0.01)
    st.markdown("**Additional info**")
    st.radio("Sampling frame and universe the same source?", ["No", "Yes"], key="same_lbl", horizontal=True)
    st.radio("Fully de-duped?", ["Yes", "No"], key="dedup_lbl", horizontal=True)
    with st.expander("Solver, split & precision"):
        solver = st.selectbox("MIP solver", [s for s in ("SCIP", "ECOS_BB") if s in cp.installed_solvers()] or ["ECOS_BB"])
        time_limit = st.number_input("SCIP time limit (s)", 1, 600, 20)
        split = st.selectbox("Split priority", ["Panel first", "Fresh first", "Proportional to frames"])
        st.slider("Max panel share of each cell (%)", 0, 100, key="pshare",
                  help="The sheet's PANEL DESIGN rule caps panel at Target/2 — i.e. 50%.")
        z = st.number_input("Z-score", value=1.644853627, format="%.6f")
        moe = st.number_input("Margin of error", value=0.075, format="%.3f")
        pp = st.number_input("p", value=0.5, format="%.2f")
        basis = st.selectbox("Population basis", ["Adjusted (auto formula)", "Uploaded adjusted sheet", "External universe"])
        ignore_field = st.checkbox("Ignore fieldwork (design from scratch)", value=False)

P = dict(total=int(st.session_state.total), min_cell=int(st.session_state.min_cell),
         max_cell=int(st.session_state.max_cell), max_bw=float(st.session_state.max_bw),
         conv=float(st.session_state.conv), pshare=st.session_state.pshare / 100,
         same=st.session_state.same_lbl == "Yes", dedup=st.session_state.dedup_lbl == "Yes",
         basis=basis if model else "Adjusted (auto formula)", split=split if model else "Panel first",
         solver=solver if model else "SCIP", time_limit=int(time_limit) if model else 20,
         z=z if model else 1.644853627, moe=moe if model else 0.075, p=pp if model else 0.5)

st.title("WBES Survey Design Studio")
if not model:
    st.info("Upload your survey-design workbook from the sidebar — the **production layout** "
            "(Inputs sheet with titled blocks, like *Survey_Design_Sri_Lanka_2025*) is read directly, "
            "parameters included. Clean one-table-per-sheet workbooks also work. Or load the demo.")
    st.stop()

IDC = model["id_cols"]
dfm = model["df"]

# dimension minimums
ninf = n_infinity(P["z"], P["moe"], P["p"])
pop_now = (dfm["uni"] if P["basis"].startswith("External") else adjusted_universe(dfm, P["same"], P["dedup"]))
default_mins = {}
for dim in IDC + ["Sector"]:
    col = dfm["Sector"] if dim == "Sector" else dfm[dim]
    for val, N in pop_now.groupby(col.astype(str), sort=False).sum().items():
        default_mins[(dim, val)] = int(round(fpc_min(N, ninf)))
sig = (st.session_state.get("_upname"), round(ninf, 3), P["basis"], P["same"], P["dedup"])
if st.session_state.get("_mins_sig") != sig:
    st.session_state.mins = dict(default_mins); st.session_state._mins_sig = sig

hd1, hd2, hd3 = st.columns([1.2, 2.5, 3])
run = hd1.button("▶ Run solver", type="primary", **FULLW)
mode_txt = ("re-design — fieldwork accounted for" if (model["has"]["completed"] and not ignore_field)
            else "fresh design")
hd2.caption(f"Mode: **{mode_txt}** · solver **{P['solver']}**")
hd3.caption(f"n={P['total']} · cell {P['min_cell']}–{P['max_cell']} · max BW {P['max_bw']:,.0f} · "
            f"conv {P['conv']:.2f} · same source **{'Yes' if P['same'] else 'No'}** · de-duped **{'Yes' if P['dedup'] else 'No'}**")

if run:
    with st.spinner(f"Solving MIP ({P['solver']})…"):
        st.session_state.res = run_solver(model, P, st.session_state.mins, ignore_field)
res = st.session_state.res

tab_in, tab_opt, tab_rev, tab_exp = st.tabs(["Inputs", "Optimization", "Review", "Export"])

# ---------------- INPUTS (mirrors the Inputs sheet) ----------------
with tab_in:
    rule = ("fresh + panel" if (P["same"] and P["dedup"]) else
            "max(universe, fresh + panel)" if P["dedup"] else "max(universe, fresh, panel)")
    adj = adjusted_universe(dfm, P["same"], P["dedup"])
    tmp = dfm.copy(); tmp["adj_auto"] = adj
    bands = [("ADJUSTED UNIVERSE", "adj_auto", "green", f"computed: {rule}"),
             ("UNIVERSE (EXTERNAL SOURCE)", "uni", "green", "")]
    bands += [("FRESH: SAMPLE FRAME", "fF", "blue", ""), ("PANEL: SAMPLE FRAME", "fP", "blue", "")]
    if model["has"]["used"]:
        bands += [("FRESH: USED CONTACTS", "uF", "amber", "any code except eligibility 14, status 2 & 4"),
                  ("PANEL: USED CONTACTS", "uP", "amber", "")]
    if model["has"]["completed"]:
        bands += [("FRESH: COMPLETED INTERVIEWS", "cF", "violet", ""), ("PANEL: COMPLETED INTERVIEWS", "cP", "violet", "")]
    if model["has"]["prev"]:
        bands += [("FRESH: MOST RECENT PREVIOUS SAMPLE DESIGN", "pF", "violet", ""),
                  ("PANEL: MOST RECENT PREVIOUS SAMPLE DESIGN", "pP", "violet", "")]
    h = 455
    for i in range(0, len(bands), 2):
        L, Rb = bands[i], bands[i + 1]
        cL, cR = st.columns(2)
        with cL:
            cap_bar(L[0], L[2], L[3]); show_plain(tmp, L[1], height=h); mini_caption(tmp.fillna(0), L[1])
        with cR:
            cap_bar(Rb[0], Rb[2], Rb[3]); show_plain(tmp.fillna(0), Rb[1], height=h); mini_caption(tmp.fillna(0), Rb[1])
    if model["has"]["completed"] and model["has"]["prev"]:
        cL, cR = st.columns(2)
        tmp["ovF"] = (tmp["cF"].fillna(0) - tmp["pF"].fillna(0)).clip(lower=0)
        tmp["ovP"] = (tmp["cP"].fillna(0) - tmp["pP"].fillna(0)).clip(lower=0)
        with cL:
            cap_bar("FRESH: OVERSHOOT TABLE", "violet", "completed − previous design, floored at 0")
            show_flag_pos(tmp, "ovF", height=h); mini_caption(tmp, "ovF")
        with cR:
            cap_bar("PANEL: OVERSHOOT TABLE", "violet", "")
            show_flag_pos(tmp, "ovP", height=h); mini_caption(tmp, "ovP")
    with st.expander("Dimension minimums (editable — defaults from z²·p(1−p)/e² with FPC)"):
        cols = st.columns(len(IDC) + 1)
        for ci, dim in enumerate(IDC + ["Sector"]):
            with cols[ci]:
                st.markdown(f"**{dim}**")
                for (d_, v_), m_ in [it for it in st.session_state.mins.items() if it[0][0] == dim]:
                    st.session_state.mins[(d_, v_)] = st.number_input(str(v_), 0, 100000, int(m_), key=f"min_{d_}_{v_}")
        sums = {}
        for (d_, _), m_ in st.session_state.mins.items(): sums[d_] = sums.get(d_, 0) + m_
        bad = [f"{d_} (Σ={s_})" for d_, s_ in sums.items() if s_ > P["total"]]
        if bad:
            st.error("Minimums exceed the sample size within: " + ", ".join(bad))
        else:
            st.success("All dimension minimums fit within the sample size.")

# ---------------- OPTIMIZATION (mirrors the Optimization sheet) ----------------
with tab_opt:
    if not res:
        st.info("Press **▶ Run solver**. This page reproduces the workbook's Optimization sheet: "
                "full design + base weights, the min/max cell bounds, the fresh/panel split, "
                "and the red re-run checks — side by side, in the same order.")
    elif not res["ok"]:
        st.error("**The design is infeasible as specified** — see the Review tab for the conflict tables "
                 "and the minimal-relaxation analysis.")
    else:
        d = res["df"]
        bwv = d.loc[d["x"] > 0, "bw"]
        redF_cells = int((d["redF"] < d["needF"]).sum()); redP_cells = int((d["redP"] < d["needP"]).sum())
        m = st.columns(6)
        m[0].metric("Sample", f"{int(d['x'].sum()):,}", f"target {P['total']:,}", delta_color="off")
        m[1].metric("Completed / to field", f"{res['achieved']:,} / {res['additional']:,}")
        m[2].metric("Fresh / panel", f"{int(d['fresh'].sum()):,} / {int(d['panel'].sum()):,}")
        m[3].metric("Max base weight", f"{bwv.max():,.0f}", f"cap {P['max_bw']:,.0f}", delta_color="off")
        m[4].metric("Red cells (F / P)", f"{redF_cells} / {redP_cells}",
                    "re-run needed" if redF_cells + redP_cells else "none",
                    delta_color="inverse" if redF_cells + redP_cells else "off")
        m[5].metric("Solver", res["solver"])
        for nmsg in res["notes"]:
            st.warning(nmsg)
        h = 455
        cL, cR = st.columns(2)
        with cL:
            cap_bar("FULL SAMPLE DESIGN", "green", "completed + additional" if res["use_field"] else "")
            show_plain(d, "x", height=h); mini_caption(d, "x")
        with cR:
            cap_bar("BASE WEIGHTS IMPLIED BY SAMPLE DESIGN", "green", "population ÷ design")
            show_heat(d, "bw", dec=1, height=h)
        cL, cR = st.columns(2)
        with cL:
            cap_bar("MINIMUM CELL SIZE", "amber",
                    "if not achievable → max(completed, expected); floor incl. weight & min-cell rules")
            show_plain(d, "lb_tot", height=h); mini_caption(d, "lb_tot")
        with cR:
            cap_bar("MAXIMUM CELL SIZE", "amber",
                    "min(completed + remaining contacts, conversion room, max cell)")
            show_plain(d, "ub_tot", height=h); mini_caption(d, "ub_tot")
        cL, cR = st.columns(2)
        with cL:
            cap_bar("FRESH SAMPLE DESIGN", "green", "")
            show_plain(d, "fresh", height=h); mini_caption(d, "fresh")
        with cR:
            cap_bar("PANEL SAMPLE DESIGN", "green",
                    f"max(panel completed, min(target×{P['pshare']:.0%}, remaining panel, …))")
            show_plain(d, "panel", height=h); mini_caption(d, "panel")
        cL, cR = st.columns(2)
        with cL:
            cap_bar("IF THERE ARE ANY IN RED — RE-RUNNING THE SOLVER IS REQUIRED", "violet",
                    "remaining FRESH contacts (frame − used); red = fewer than still needed")
            show_red(d, "redF", "needF", height=h)
        with cR:
            cap_bar("IF THERE ARE ANY IN RED — RE-RUNNING THE SOLVER IS REQUIRED", "violet",
                    "remaining PANEL contacts; red = fewer than still needed")
            show_red(d, "redP", "needP", height=h)
        if "overF" in d:
            cL, cR = st.columns(2)
            with cL:
                cap_bar("FRESH: OVERSHOOT vs PREVIOUS DESIGN", "violet", "red = completed above plan")
                show_flag_pos(d, "overF", height=h)
            with cR:
                cap_bar("PANEL: OVERSHOOT vs PREVIOUS DESIGN", "violet", "")
                show_flag_pos(d, "overP", height=h)

# ---------------- REVIEW ----------------
with tab_rev:
    if not res:
        st.info("Run the solver first.")
    elif not res["ok"]:
        c = res["conflicts"]
        if c["overall"]:
            st.subheader("Overall conflicts")
            st.dataframe(pd.DataFrame(c["overall"], columns=["Cause", "Required", "Capacity"]), hide_index=True)
        if c["dims"]:
            st.subheader("Dimension conflicts")
            st.dataframe(pd.DataFrame(c["dims"], columns=["Constraint", "Min", "Feasible max"]), hide_index=True)
        if c["cells"]:
            st.subheader("Cell conflicts (lower bound > feasible max)")
            st.dataframe(res["df"].iloc[c["cells"]][IDC + ["Sector", "pop", "fF", "fP", "lb_tot", "ub_tot"]],
                         hide_index=True)
        st.subheader("Minimal relaxation needed (slack LP)")
        st.dataframe(res["slack"], hide_index=True, **FULLW)
        st.markdown("**Repair guidance:** lower the sample size or raise the conversion rate if capacity is short; "
                    "raise the max base weight or lower the min cell size for cell conflicts; "
                    "relax the listed dimension minimums otherwise.")
    else:
        d = res["df"]
        st.subheader("Dimension minimum checks")
        rows = []
        for (dim, val), mreq in st.session_state.mins.items():
            col = d["Sector"] if dim == "Sector" else d[dim]
            got = int(d.loc[col.astype(str) == str(val), "x"].sum())
            rows.append({"Dimension": dim, "Value": val, "Min": mreq, "Allocated": got,
                         "Status": "met" if got >= mreq else f"short {mreq - got}"})
        st.dataframe(pd.DataFrame(rows), hide_index=True, **FULLW)
        st.subheader("Deviation from proportional (allocated − target)")
        d2 = d.copy(); d2["dev"] = d2["x"] - d2["t"]
        st.dataframe(pivot(d2, "dev", dec=1).style.format("{:,.1f}", na_rep="–")
                     .background_gradient(cmap="PuOr", axis=None, vmin=-float(np.abs(d2['dev']).max() or 1),
                                          vmax=float(np.abs(d2['dev']).max() or 1)),
                     **FULLW, height=455)
        st.metric("Sum of squared deviations", f"{float(((d['x'] - d['t']) ** 2).sum()):,.1f}")

# ---------------- EXPORT ----------------
with tab_exp:
    st.download_button("⬇ Excel workbook with all result tables",
                       data=to_excel(model, P, st.session_state.mins, res) if (res and res["ok"]) else b"",
                       file_name="WBES_design_results.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       disabled=not (res and res.get("ok")))
    st.caption("Sheets: Parameters · Dimension minimums · Full / Fresh / Panel design · Base weights · "
               "Min/Max cell size · Remaining contacts (red checks) · Overshoots · per-cell detail.")
