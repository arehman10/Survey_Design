import streamlit as st
import pandas as pd
import numpy as np
import math
import cvxpy as cp
import io
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule
import matplotlib.colors as mcolors
from matplotlib.colors import LinearSegmentedColormap, TwoSlopeNorm
import base64
import textwrap, html
import swiglpk as glpk  # If needed for GLPK
from pandas.io.formats.style import Styler
import uuid
import json
import os

###############################################################################
# 1) HELPER FUNCTIONS
###############################################################################


SESSIONS_DIR = "sessions"  # local folder to store session JSON files


def show_side_by_side(df_left, title_left, df_right, title_right, key_left, key_right, editable=False):
    """Render two dataframes side-by-side; use_editor=True to allow editing."""
    c1, c2 = st.columns(2, gap="small")
    with c1:
        st.markdown(f"**{title_left}**")
        if editable:
            st.data_editor(df_left, use_container_width=True, height=520, key=key_left)
        else:
            st.dataframe(df_left, use_container_width=True, height=520, key=key_left)
    with c2:
        st.markdown(f"**{title_right}**")
        if editable:
            st.data_editor(df_right, use_container_width=True, height=520, key=key_right)
        else:
            st.dataframe(df_right, use_container_width=True, height=520, key=key_right)

def init_sessions_dir():
    """Ensure we have a 'sessions' subfolder to store JSON files."""
    if not os.path.exists(SESSIONS_DIR):
        os.makedirs(SESSIONS_DIR)

def save_session_to_file(session_id, data_dict):
    """Write data_dict to a JSON file in SESSIONS_DIR using session_id as filename."""
    init_sessions_dir()
    file_path = os.path.join(SESSIONS_DIR, f"{session_id}.json")
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data_dict, f, ensure_ascii=False, indent=2)

def load_session_from_file(session_id):
    """Read the JSON from SESSIONS_DIR. Return dict or None if not found."""
    file_path = os.path.join(SESSIONS_DIR, f"{session_id}.json")
    if not os.path.exists(file_path):
        return None
    with open(file_path, "r", encoding="utf-8") as f:
        return json.load(f)

def add_grand_total_row(df, key_col_name):
    """Return df with an extra row that holds the column totals."""
    grand = df[["PanelSample", "FreshSample", "SampleTotal"]].sum().to_frame().T
    grand[key_col_name] = "Grand Total"
    return pd.concat([df, grand], ignore_index=True)[df.columns]

HTML_TEMPLATE = textwrap.dedent("""\
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="utf-8">
      <title>{page_title}</title>
      <style>
        body      {{ font-family: Arial, sans-serif; margin: 2rem; }}
        h2        {{ color:#154360; }}
        table     {{ border-collapse: collapse; margin-bottom: 1.5rem; width:100%; }}
        th, td    {{ border: 1px solid #bbb; padding: 4px 6px; font-size: 14px; }}
        th        {{ background:#eef; }}
        tr:nth-child(even) td {{ background:#f9f9f9; }}
        .scrollbox {{ max-height:400px; overflow:auto; border:1px solid #ccc; padding:4px; }}
      </style>
    </head>
    <body>
      <h1>{page_title}</h1>
      {body_html}
    </body>
    </html>
""")

def dfs_to_html(sections, page_title="Snapshot"):
    """
    sections = [(title, dataframe_or_styler), ...]
    Returns bytes of a single HTML file with scrollable <div> around each table.
    """
    parts = []
    for title, obj in sections:
        parts.append(f"<h2>{html.escape(title)}</h2>")
        parts.append('<div class="scrollbox">')
        if isinstance(obj, Styler):
            # Keep styling if it's a Styler object
            parts.append(obj.to_html())
        else:
            parts.append(obj.to_html(index=False, border=0, justify="center"))
        parts.append("</div>")
    full = HTML_TEMPLATE.format(body_html="\n".join(parts), page_title=html.escape(page_title))
    return full.encode("utf-8")

def compute_n_infinity(z_score, margin_of_error, p):
    return (z_score ** 2) * p * (1 - p) / (margin_of_error ** 2)

def compute_fpc_min(N, n_infinity):
    """
    Finite Population Correction-based recommended minimum sample for a group of size N.
    """
    if N <= 0:
        return 0
    else:
        return n_infinity / (1 + n_infinity / N)

def create_combined_table_with_totals(df_long):
    """
    Pivot table for (OptimizedSample, BaseWeight) with row/col totals.
    """
    pivot_sample = pd.pivot_table(
        df_long,
        index=["Region", "Size"],
        columns="Industry",
        values="OptimizedSample",
        aggfunc='sum',
        fill_value=0,
        margins=True,
        margins_name="GrandTotal",
        sort=False
    )

    pivot_bw = pd.pivot_table(
        df_long,
        index=["Region", "Size"],
        columns="Industry",
        values="BaseWeight",
        aggfunc='sum',
        fill_value=0,
        margins=True,
        margins_name="GrandTotal",
        sort=False
    )

    all_inds = set(pivot_sample.columns).union(pivot_bw.columns)
    combined = pd.DataFrame(index=pivot_sample.index)
    for ind in all_inds:
        sample_col = f"{ind}_Sample"
        bw_col     = f"{ind}_BaseWeight"
        combined[sample_col] = pivot_sample[ind] if ind in pivot_sample.columns else np.nan
        combined[bw_col]     = pivot_bw[ind]     if ind in pivot_bw.columns else np.nan

    # reorder so that "GrandTotal_Sample" & "GrandTotal_BaseWeight" appear first
    col_list = list(combined.columns)
    if "GrandTotal_Sample" in col_list:
        col_list.remove("GrandTotal_Sample")
        col_list.insert(0, "GrandTotal_Sample")
    if "GrandTotal_BaseWeight" in col_list:
        col_list.remove("GrandTotal_BaseWeight")
        col_list.insert(0, "GrandTotal_BaseWeight")

    combined = combined[col_list]
    return combined

def write_scenario_sheets_to_writer(
    writer,
    df_combined,
    pivot_population,
    pivot_propsample,
    scenario_label=""
):
    """
    Write scenario's 4 sheets directly to the *same* Excel writer workbook:
      1) <scenario_label>Allocated Sample
      2) <scenario_label>Proportional Sample
      3) <scenario_label>Sample_with_baseweight (with color scale)
      4) <scenario_label>Population

    This avoids the error "Cannot copy between worksheets from different workbooks."
    """
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import ColorScaleRule

    df_out = df_combined.reset_index(drop=False)
    n_rows = df_out.shape[0]

    # We separate columns
    id_cols     = [c for c in ["Region","Size"] if c in df_out.columns]
    sample_cols = [c for c in df_out.columns if c.endswith("_Sample")]
    bw_cols     = [c for c in df_out.columns if c.endswith("_BaseWeight")]
    df_out = df_out[id_cols + sample_cols + bw_cols]

    # For color scaling
    norm_bw_cols = [c for c in bw_cols if c!="GrandTotal_BaseWeight"]
    if len(df_out) > 1 and norm_bw_cols:
        norm_df = df_out.iloc[:-1]
        global_min = norm_df[norm_bw_cols].min().min()
        global_max = norm_df[norm_bw_cols].max().max()
        global_mid = np.percentile(norm_df[norm_bw_cols].stack(), 50)
    else:
        global_min=0
        global_mid=0
        global_max=0

    # 1) Allocated Sample
    sheet_samples = f"{scenario_label}Allocated Sample" if scenario_label else "Allocated Sample"
    df_samples_only = df_out[id_cols + sample_cols].copy()
    df_samples_only.to_excel(writer, sheet_name=sheet_samples, index=False)
    
    # 2) Proportional Sample
    sheet_prop = f"{scenario_label}Proportional Sample" if scenario_label else "Proportional Sample"
    pivot_propsample_rounded = pivot_propsample.round(0).astype(int)
    pivot_propsample_rounded.reset_index().to_excel(
        writer, sheet_name=sheet_prop, index=False
    )
    ws_prop = writer.sheets[sheet_prop]
    first_row = 2
    last_row  = ws_prop.max_row
    first_col = 3
    last_col  = ws_prop.max_column
    for row in ws_prop.iter_rows(min_row=first_row, max_row=last_row,
                                 min_col=first_col, max_col=last_col):
        for cell in row:
            cell.number_format = "0"

    # 3) Sample_with_baseweight
    sheet_sw = f"{scenario_label}Sample_with_baseweight" if scenario_label else "Sample_with_baseweight"
    df_out.to_excel(writer, sheet_name=sheet_sw, index=False)
    ws_sw = writer.sheets[sheet_sw]

    def make_rule():
        return ColorScaleRule(
            start_type="num", start_value=global_min, start_color="00FF00",
            mid_type="num",   mid_value=global_mid,  mid_color="FFFF00",
            end_type="num",   end_value=global_max,  end_color="FF0000",
        )

    for col_name in norm_bw_cols:
        col_idx = df_out.columns.get_loc(col_name)+1  # +1 for 1-based
        excel_col = get_column_letter(col_idx)
        rng       = f"{excel_col}2:{excel_col}{n_rows}"
        ws_sw.conditional_formatting.add(rng, make_rule())
        # format to 0.0
        for cell in ws_sw[f"{excel_col}2":f"{excel_col}{n_rows}"]:
            cell[0].number_format = "0.0"

    # 4) Population
    sheet_pop = f"{scenario_label}Population" if scenario_label else "Population"
    pivot_population.to_excel(writer, sheet_name=sheet_pop, index=True)


from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

def _autofit_columns(ws, min_width=9, max_width=38):
    """Auto-fit all used columns based on cell string length."""
    dims = {}
    for row in ws.rows:
        for cell in row:
            if cell.value is None:
                continue
            col = cell.column_letter
            val = str(cell.value)
            dims[col] = max(dims.get(col, 0), len(val))
    for col, width in dims.items():
        ws.column_dimensions[col].width = max(min_width, min(width + 2, max_width))

def _format_int_block(ws, top_row, left_col, n_rows, n_cols):
    """Apply integer formatting 0 to a rectangular block."""
    for r in ws.iter_rows(min_row=top_row, max_row=top_row+n_rows-1,
                          min_col=left_col, max_col=left_col+n_cols-1):
        for c in r:
            c.number_format = "0"

def write_panel_fresh_and_mins_sheet(
    writer,
    sheet_name,
    panel_df,
    fresh_df,
    region_mins_df,
    size_mins_df,
    industry_mins_df,
    title_left="Panel",
    title_right="Fresh",
):
    """
    Create one sheet with Panel (left) and Fresh (right) side-by-side,
    and Region/Size/Industry minimums below (three blocks side-by-side).
    """
    gap_cols = 2
    startrow = 2   # leave top row for big title row, row 2 for subheaders
    startcol = 1   # 1-based for openpyxl, but pandas uses 0-based -> we pass startcol-1

    # 1) Write the two pivots
    panel_df = panel_df.copy()
    fresh_df = fresh_df.copy()
    panel_df.to_excel(writer, sheet_name=sheet_name, index=False,
                      startrow=startrow, startcol=startcol-1)
    fresh_startcol = startcol + panel_df.shape[1] + gap_cols
    fresh_df.to_excel(writer, sheet_name=sheet_name, index=False,
                      startrow=startrow, startcol=fresh_startcol-1)

    ws = writer.sheets[sheet_name]

    # 2) Titles above each table
    big = Font(size=13, bold=True)
    smallbold = Font(bold=True)
    ws.cell(row=1, column=startcol, value=f"{title_left}").font = big
    ws.cell(row=1, column=fresh_startcol, value=f"{title_right}").font = big

    # 3) Freeze panes just below headers
    ws.freeze_panes = ws.cell(row=startrow+1, column=startcol)

    # 4) Integer formatting for numeric parts of both pivots (all columns after Region/Size)
    # Detect numeric columns by dtype in the dataframes
    nrows_panel = panel_df.shape[0] + 1   # + header
    ncols_panel = panel_df.shape[1]
    nrows_fresh = fresh_df.shape[0] + 1
    ncols_fresh = fresh_df.shape[1]

    # Panel numeric block: assume first two columns are Region/Size (adjust if needed)
    if ncols_panel >= 3:
        _format_int_block(ws,
                          top_row=startrow+1,
                          left_col=startcol+2,
                          n_rows=nrows_panel-1,
                          n_cols=ncols_panel-2)

    # Fresh numeric block
    if ncols_fresh >= 3:
        _format_int_block(ws,
                          top_row=startrow+1,
                          left_col=fresh_startcol+2,
                          n_rows=nrows_fresh-1,
                          n_cols=ncols_fresh-2)

    # 5) Dimension minimums (three blocks below the lower of the two tables)
    below_row = startrow + max(nrows_panel, nrows_fresh) + 2
    ws.cell(row=below_row-1, column=startcol,
            value="Dimension Minimums").font = smallbold

    reg_startcol  = startcol
    size_startcol = reg_startcol  + region_mins_df.shape[1] + gap_cols
    ind_startcol  = size_startcol + size_mins_df.shape[1]   + gap_cols

    # Normalize column names and order
    reg_df  = region_mins_df.copy()
    size_df = size_mins_df.copy()
    ind_df  = industry_mins_df.copy()

    # Write the three blocks
    reg_df.to_excel(writer, sheet_name=sheet_name, index=False,
                    startrow=below_row, startcol=reg_startcol-1)
    size_df.to_excel(writer, sheet_name=sheet_name, index=False,
                     startrow=below_row, startcol=size_startcol-1)
    ind_df.to_excel(writer, sheet_name=sheet_name, index=False,
                    startrow=below_row, startcol=ind_startcol-1)

    # Integer formatting on MinNeeded columns if present
    # Find "MinNeeded" col index per df (1-based in Excel)
    def _format_mins_block(df, left_col):
        if "MinNeeded" in df.columns:
            col_offset = list(df.columns).index("MinNeeded")
            _format_int_block(ws,
                              top_row=below_row+1,
                              left_col=left_col + col_offset,
                              n_rows=df.shape[0],
                              n_cols=1)

    _format_mins_block(reg_df,  reg_startcol)
    _format_mins_block(size_df, size_startcol)
    _format_mins_block(ind_df,  ind_startcol)

    # 6) Pretty widths
    _autofit_columns(ws)

    # 7) Light header fill for readability (panel + fresh + the three mins blocks)
    header_fill = PatternFill(start_color="FFEFEFEF", end_color="FFEFEFEF", fill_type="solid")
    # Panel header row
    for j in range(ncols_panel):
        ws.cell(row=startrow, column=startcol+j).fill = header_fill
        ws.cell(row=startrow, column=startcol+j).font = smallbold
    # Fresh header row
    for j in range(ncols_fresh):
        ws.cell(row=startrow, column=fresh_startcol+j).fill = header_fill
        ws.cell(row=startrow, column=fresh_startcol+j).font = smallbold
    # Mins headers
    for j in range(reg_df.shape[1]):
        ws.cell(row=below_row, column=reg_startcol+j).fill = header_fill
        ws.cell(row=below_row, column=reg_startcol+j).font = smallbold
    for j in range(size_df.shape[1]):
        ws.cell(row=below_row, column=size_startcol+j).fill = header_fill
        ws.cell(row=below_row, column=size_startcol+j).font = smallbold
    for j in range(ind_df.shape[1]):
        ws.cell(row=below_row, column=ind_startcol+j).fill = header_fill
        ws.cell(row=below_row, column=ind_startcol+j).font = smallbold

###############################################################################
# 2) FEASIBILITY & SLACK DIAGNOSTICS
###############################################################################
def detailed_feasibility_check(df_long,
                               total_sample,
                               min_cell_size,
                               max_cell_size,
                               max_base_weight,
                               conversion_rate,
                               dimension_mins):

    n_cells = len(df_long)
    feasible_max_by_i, cell_conflicts = [], []
    dimension_conflicts, overall_conflicts = [], []

    for i in range(n_cells):
        pop_i   = float(df_long.loc[i, "Population"])
        panel_i = float(df_long.loc[i, "PanelPop"]) if "PanelPop" in df_long.columns else 0.0
        fresh_i = float(df_long.loc[i, "FreshPop"]) if "FreshPop" in df_long.columns else 0.0
        avail_i = panel_i + fresh_i if ("PanelPop" in df_long.columns and "FreshPop" in df_long.columns) else pop_i

        if avail_i <= 0:
            fmax = 0
        else:
            conv_ub = math.ceil(avail_i * conversion_rate)
            fmax = min(avail_i, max_cell_size, conv_ub)

        lb_bw = (pop_i / max_base_weight) if max_base_weight > 0 else 0
        if fmax >= min_cell_size and avail_i > 0:
            lb = max(lb_bw, min_cell_size, 0)
        else:
            lb = max(lb_bw, 0)

        if lb > fmax + 1e-9:
            cell_conflicts.append({
                "CellIndex": i,
                "Region": df_long.loc[i, "Region"],
                "Size": df_long.loc[i, "Size"],
                "Industry": df_long.loc[i, "Industry"],
                "Population": pop_i,
                "PanelPop": panel_i,
                "FreshPop": fresh_i,
                "LowerBound": lb,
                "FeasibleMax": fmax,
                "ShortBy": lb - fmax,
                "Reason": "Cell min > feasible max (given panel+fresh capacity)"
            })

        feasible_max_by_i.append(fmax)

    sum_feas = sum(feasible_max_by_i)
    if sum_feas < total_sample:
        overall_conflicts.append({
            "TotalSample": total_sample,
            "SumFeasibleMax": sum_feas,
            "ShortBy": total_sample - sum_feas,
            "Reason": "Overall capacity (panel+fresh) < total_sample"
        })

    # dimension checks
    dim_idx = {"Region": {}, "Size": {}, "Industry": {}}
    for dt in dim_idx:
        for val in df_long[dt].unique():
            idx_list = df_long.index[df_long[dt] == val].tolist()
            dim_idx[dt][val] = idx_list

    for dt, val_dict in dimension_mins.items():
        for val, req_min in val_dict.items():
            if req_min > 0:
                idx_list = dim_idx[dt].get(val, [])
                sum_fmax_dim = sum(feasible_max_by_i[i] for i in idx_list)
                if sum_fmax_dim < req_min:
                    dimension_conflicts.append({
                        "DimType": dt,
                        "DimName": val,
                        "RequiredMin": req_min,
                        "SumFeasibleMax": sum_fmax_dim,
                        "ShortBy": req_min - sum_fmax_dim,
                        "Reason": "Dimension min exceeds feasible capacity (panel+fresh)"
                    })

    return (
        pd.DataFrame(overall_conflicts),
        pd.DataFrame(cell_conflicts),
        pd.DataFrame(dimension_conflicts)
    )


def diagnose_infeasibility_slacks(df_long,
                                  total_sample,
                                  min_cell_size,
                                  max_cell_size,
                                  max_base_weight,
                                  dimension_mins,
                                  conversion_rate):
    """
    Slack diagnostic using panel+fresh capacity when available.
    """
    import cvxpy as cp
    n = len(df_long)
    x = cp.Variable(n, nonneg=True)
    s_tot = cp.Variable(nonneg=True)
    constraints = [cp.sum(x) + s_tot == total_sample]

    # dimension index
    dim_idx = {"Region": {}, "Size": {}, "Industry": {}}
    for dt in dim_idx:
        for val in df_long[dt].unique():
            idx_list = df_long.index[df_long[dt] == val].tolist()
            dim_idx[dt][val] = idx_list

    dim_slacks = {}
    for dt, val_dict in dimension_mins.items():
        for val, req_min in val_dict.items():
            if req_min > 0:
                sd = cp.Variable(nonneg=True)
                dim_slacks[(dt, val)] = sd
                idx_list = dim_idx[dt][val]
                constraints.append(cp.sum(x[idx_list]) + sd >= req_min)

    cell_slacks = []
    for i in range(n):
        pop_i   = float(df_long.loc[i, "Population"])
        panel_i = float(df_long.loc[i, "PanelPop"]) if "PanelPop" in df_long.columns else 0.0
        fresh_i = float(df_long.loc[i, "FreshPop"]) if "FreshPop" in df_long.columns else 0.0
        avail_i = panel_i + fresh_i if ("PanelPop" in df_long.columns and "FreshPop" in df_long.columns) else pop_i

        if avail_i <= 0:
            constraints.append(x[i] == 0)
            continue

        conv_ub = math.ceil(avail_i * conversion_rate)
        fmax = min(avail_i, max_cell_size, conv_ub)
        lb_bw = (pop_i / max_base_weight) if max_base_weight > 0 else 0
        if fmax >= min_cell_size and avail_i > 0:
            lb = max(lb_bw, min_cell_size, 0)
        else:
            lb = max(lb_bw, 0)

        s_cell = cp.Variable(nonneg=True)
        cell_slacks.append(s_cell)
        constraints += [x[i] + s_cell >= lb, x[i] <= fmax]

    slack_vars = [s_tot] + list(dim_slacks.values()) + cell_slacks
    obj = cp.Minimize(cp.sum(slack_vars))
    prob = cp.Problem(obj, constraints)
    try:
        prob.solve(solver="ECOS", verbose=False)
    except Exception as e:
        return None, None, f"Diagnostic solver error: {e}"

    if prob.status in ["infeasible", "unbounded"]:
        return None, None, f"Diagnostic problem is {prob.status}."

    x_sol = x.value
    df_slack_sol = df_long.copy()
    df_slack_sol["SlackSolution_x"] = x_sol

    slack_usage = []
    if s_tot.value > 1e-8:
        slack_usage.append({"Constraint": "TotalSample", "SlackUsed": s_tot.value})
    for (dt, val), var in dim_slacks.items():
        if var.value > 1e-8:
            slack_usage.append({"Constraint": f"DimensionMin {dt}={val}", "SlackUsed": var.value})
    for i, var in enumerate(cell_slacks):
        if var.value > 1e-8:
            r = df_long.loc[i, "Region"]; s = df_long.loc[i, "Size"]; ind = df_long.loc[i, "Industry"]
            slack_usage.append({"Constraint": f"CellMin i={i} (R={r}, Sz={s}, Ind={ind})",
                                "SlackUsed": var.value})

    df_slack_usage = pd.DataFrame(slack_usage).sort_values("SlackUsed", ascending=False)
    return df_slack_sol, df_slack_usage, prob.status

###############################################################################
# 3) MAIN SOLVER
###############################################################################
def run_optimization(df_wide,
                     total_sample,
                     min_cell_size,
                     max_cell_size,
                     max_base_weight,
                     solver_choice,
                     dimension_mins,
                     conversion_rate,
                     df_frame_panel_wide,
                     df_frame_fresh_wide,
                     time_limit_sec=20,
                     fallback_relax_and_round=True,
                     per_source_conv_caps=False  # set True if you want per-source conversion caps
                     ):
    """
    Fast MIP on totals x[i] using *frame* capacities (panel/fresh) as feasibility
    and then a guaranteed-feasible split into PanelAllocated/FreshAllocated.

    - Universe (Population, PropSample) comes from df_wide (adjusted with panel/fresh as you already do).
    - Capacity/feasibility come from frame tables:
        PanelAllocated[i] <= frame_panel
        FreshAllocated[i] <= frame_fresh
        x[i]              <= frame_panel + frame_fresh (and conversion & max_cell_size)
    - Dimension minimums apply to totals x (not per-source).
    """
    import cvxpy as cp

    # ---------- Build long tables ----------
    id_cols  = ["Region", "Size"]
    data_cols = [c for c in df_wide.columns if c not in id_cols]

    # Population/universe
    df_pop = df_wide.melt(
        id_vars=id_cols, value_vars=data_cols,
        var_name="Industry", value_name="Population"
    ).fillna({"Population": 0})

    # Capacities from FRAME sheets (not the 'panel'/'fresh' universes)
    df_panel = df_frame_panel_wide.melt(
        id_vars=id_cols, value_vars=[c for c in df_frame_panel_wide.columns if c not in id_cols],
        var_name="Industry", value_name="PanelPop"
    )
    df_fresh = df_frame_fresh_wide.melt(
        id_vars=id_cols, value_vars=[c for c in df_frame_fresh_wide.columns if c not in id_cols],
        var_name="Industry", value_name="FreshPop"
    )

    df_long = (df_pop
               .merge(df_panel, on=["Region","Size","Industry"], how="left")
               .merge(df_fresh, on=["Region","Size","Industry"], how="left"))

    df_long[["PanelPop","FreshPop"]] = df_long[["PanelPop","FreshPop"]].fillna(0.0)
    df_long["PanelPop"] = df_long["PanelPop"].astype(float)
    df_long["FreshPop"] = df_long["FreshPop"].astype(float)
    df_long["Avail"]    = df_long["PanelPop"] + df_long["FreshPop"]

    # Proportional target for objective
    total_pop = float(df_long["Population"].sum())
    df_long["PropSample"] = (df_long["Population"] * (total_sample / total_pop)) if total_pop > 0 else 0.0

    # ---------- Single-constraint feasibility pre-check (capacity-aware) ----------
    df_overall, df_cells, df_dims = detailed_feasibility_check(
        df_long, total_sample, min_cell_size, max_cell_size, max_base_weight,
        conversion_rate, dimension_mins
    )
    if (not df_overall.empty) or (not df_cells.empty) or (not df_dims.empty):
        # Return None to signal conflict, same as your app expects
        return None, df_cells, df_dims, df_overall

    # ---------- Build MIP on totals x ----------
    n = len(df_long)
    x = cp.Variable(n, integer=True)

    avail   = df_long["Avail"].to_numpy()
    pop_arr = df_long["Population"].to_numpy()

    # Upper bounds: availability, conversion, cell max
    if per_source_conv_caps:
        # Optional: cap each source by its own conversion, then sum caps
        panel_cap_eff = np.minimum(df_long["PanelPop"].to_numpy(),
                                   np.ceil(df_long["PanelPop"].to_numpy() * conversion_rate))
        fresh_cap_eff = np.minimum(df_long["FreshPop"].to_numpy(),
                                   np.ceil(df_long["FreshPop"].to_numpy() * conversion_rate))
        conv_ub = panel_cap_eff + fresh_cap_eff
    else:
        conv_ub = np.ceil(avail * conversion_rate)

    ub_float = np.minimum.reduce([avail, conv_ub, np.full(n, max_cell_size, dtype=float)])
    ub_int   = np.maximum(0, np.floor(ub_float)).astype(int)

    # Lower bounds: base-weight and min cell size when feasible
    lb_bw = (pop_arr / max_base_weight) if max_base_weight > 0 else np.zeros(n)
    min_cell_arr = np.full(n, min_cell_size, dtype=float)
    lb_float = np.where((ub_float >= min_cell_size) & (avail > 0),
                        np.maximum.reduce([lb_bw, min_cell_arr, np.zeros(n)]),
                        np.maximum(lb_bw, 0))
    lb_int = np.maximum(0, np.ceil(lb_float)).astype(int)

    # If any lb > ub after rounding, we already caught that in feasibility_check; but guard anyway:
    infeas_mask = lb_int > ub_int
    if np.any(infeas_mask):
        # translate to the app's conflict shape
        bad_idx = np.where(infeas_mask)[0]
        _cells = []
        for i in bad_idx:
            _cells.append({
                "CellIndex": int(i),
                "Region": df_long.loc[i,"Region"],
                "Size": df_long.loc[i,"Size"],
                "Industry": df_long.loc[i,"Industry"],
                "LowerBound": int(lb_int[i]),
                "FeasibleMax": int(ub_int[i]),
                "Reason": "Rounded LB > UB"
            })
        return None, pd.DataFrame(_cells), pd.DataFrame(), pd.DataFrame()

    # Constraints
    constraints = [x >= lb_int, x <= ub_int, cp.sum(x) == total_sample]

    # Dimension mins on totals x
    dim_idx = {"Region": {}, "Size": {}, "Industry": {}}
    for dt in dim_idx:
        for val in df_long[dt].unique():
            dim_idx[dt][val] = df_long.index[df_long[dt] == val].tolist()
    for dt, val_dict in dimension_mins.items():
        for val, req_min in val_dict.items():
            if req_min > 0 and val in dim_idx[dt]:
                constraints.append(cp.sum(x[dim_idx[dt][val]]) >= req_min)

    # Objective
    target = df_long["PropSample"].to_numpy()
    objective = cp.Minimize(cp.sum_squares(x - target))
    prob = cp.Problem(objective, constraints)

    # ---------- Solve with time limits; then relax+round if needed ----------
    solver_used = None
    x_val = None
    solver_kwargs = dict(warm_start=True, verbose=False)
    if solver_choice in ["SCIP", "ECOS_BB"]:
        chosen = [solver_choice] + [s for s in ["SCIP","ECOS_BB"] if s != solver_choice]
    else:
        chosen = ["SCIP", "ECOS_BB"]

    solution_found, last_error = False, None
    for s in chosen:
        try:
            if s == "SCIP":
                solver_kwargs["scip_params"] = {"limits/time": float(time_limit_sec)}
                _ = prob.solve(solver=cp.SCIP, **solver_kwargs)
            else:
                _ = prob.solve(solver=cp.ECOS_BB, mi_max_iters=100000, **solver_kwargs)

            if prob.status not in ["infeasible", "unbounded"] and x.value is not None:
                xr = np.rint(np.asarray(x.value).reshape(-1)).astype(int)
                # clamp to integer bounds
                xr = np.minimum(np.maximum(xr, lb_int), ub_int)

                # fix sum if needed
                delta = int(total_sample - xr.sum())
                if delta != 0:
                    if delta > 0:
                        room = (ub_int - xr)
                        for idx in np.argsort(-room):
                            if delta == 0: break
                            add = min(delta, int(room[idx]))
                            if add > 0:
                                xr[idx] += add
                                delta -= add
                    else:
                        room = (xr - lb_int)
                        for idx in np.argsort(-room):
                            if delta == 0: break
                            sub = min(-delta, int(room[idx]))
                            if sub > 0:
                                xr[idx] -= sub
                                delta += sub

                if xr.sum() == total_sample:
                    x_val = xr
                    solver_used = s
                    solution_found = True
                    break
        except Exception as e:
            last_error = e

    if not solution_found and fallback_relax_and_round:
        try:
            x_rel = cp.Variable(n)
            prob_rel = cp.Problem(
                cp.Minimize(cp.sum_squares(x_rel - target)),
                [x_rel >= lb_int, x_rel <= ub_int, cp.sum(x_rel) == total_sample]
            )
            _ = prob_rel.solve(solver=cp.OSQP, warm_start=True, verbose=False)
            xr = np.clip(np.rint(x_rel.value), lb_int, ub_int).astype(int)

            # fix sum
            delta = int(total_sample - xr.sum())
            if delta != 0:
                if delta > 0:
                    room = (ub_int - xr)
                    for idx in np.argsort(-room):
                        if delta == 0: break
                        add = min(delta, int(room[idx]))
                        if add > 0:
                            xr[idx] += add
                            delta -= add
                else:
                    room = (xr - lb_int)
                    for idx in np.argsort(-room):
                        if delta == 0: break
                        sub = min(-delta, int(room[idx]))
                        if sub > 0:
                            xr[idx] -= sub
                            delta += sub

            # ensure dimension mins (light-touch top-up inside capacity)
            for dt, val_dict in dimension_mins.items():
                for val, req_min in val_dict.items():
                    if req_min > 0:
                        idxs = dim_idx[dt].get(val, [])
                        short = int(req_min - xr[idxs].sum())
                        if short > 0:
                            slack = (ub_int[idxs] - xr[idxs])
                            for j in np.argsort(-slack):
                                if short == 0: break
                                jj = idxs[j]
                                add = min(short, int(slack[j]))
                                if add > 0:
                                    xr[jj] += add
                                    short -= add

            if xr.sum() != total_sample:
                raise RuntimeError("Rounding failed to hit total sample exactly.")

            x_val = xr
            solver_used = "RELAX+ROUND"
            solution_found = True
        except Exception as e:
            last_error = e

    if not solution_found:
        raise ValueError(f"No solver found a feasible solution (last error: {last_error})")

    # ---------- Feasible per-source split (never exceeds frame capacities) ----------
    out = df_long.copy()
    out["OptimizedSample"] = x_val.astype(int)

    panel_cap = df_long["PanelPop"].to_numpy().astype(int)
    fresh_cap = df_long["FreshPop"].to_numpy().astype(int)

    if per_source_conv_caps:
        panel_cap_eff = np.minimum(panel_cap, np.ceil(panel_cap * conversion_rate).astype(int))
        fresh_cap_eff = np.minimum(fresh_cap, np.ceil(fresh_cap * conversion_rate).astype(int))
    else:
        panel_cap_eff = panel_cap
        fresh_cap_eff = fresh_cap

    x_arr = out["OptimizedSample"].to_numpy()
    p_alloc = np.zeros_like(x_arr, dtype=int)
    f_alloc = np.zeros_like(x_arr, dtype=int)

    for i in range(len(x_arr)):
        xi = int(x_arr[i])
        pc = int(panel_cap_eff[i])
        fc = int(fresh_cap_eff[i])

        if xi <= 0 or (pc <= 0 and fc <= 0):
            continue

        # start with equal split
        p = min(xi // 2, pc)
        f = min(xi - p, fc)

        # allocate remainder to the side with room (prefer fresh first)
        rem = xi - (p + f)
        if rem > 0:
            f_take = min(rem, max(0, fc - f)); f += f_take; rem -= f_take
        if rem > 0:
            p_take = min(rem, max(0, pc - p)); p += p_take; rem -= p_take

        # At this point, because xi <= pc+fc, rem should be 0
        if rem > 0:
            # Extremely defensive: this should never happen if ub_int respected Avail
            raise ValueError(f"Split infeasible at i={i}: xi={xi}, panel_cap={pc}, fresh_cap={fc}")

        p_alloc[i] = p
        f_alloc[i] = f

    out["PanelAllocated"] = p_alloc
    out["FreshAllocated"] = f_alloc

    # Sanity checks (feasible by construction)
    if np.any(out["PanelAllocated"].to_numpy() > panel_cap_eff) or \
       np.any(out["FreshAllocated"].to_numpy() > fresh_cap_eff) or \
       np.any((out["PanelAllocated"] + out["FreshAllocated"]).to_numpy() != out["OptimizedSample"].to_numpy()):
        raise ValueError("Internal allocation error: per-source caps violated or totals mismatch.")

    # Base weights
    def basew(row):
        return (row["Population"] / row["OptimizedSample"]) if row["OptimizedSample"] > 0 else np.nan
    out["BaseWeight"] = out.apply(basew, axis=1)

    return out, None, None, solver_used



###############################################################################
# 4) ALLOCATE BETWEEN PANEL & FRESH
###############################################################################
def allocate_panel_fresh(df_long_sol, df_panel_wide, df_fresh_wide):
    identifier_cols = ["Region","Size"]
    inds = [c for c in df_panel_wide.columns if c not in identifier_cols]

    panel_long = df_panel_wide.melt(id_vars=identifier_cols, value_vars=inds,
                                    var_name="Industry", value_name="PanelPop")
    fresh_long = df_fresh_wide.melt(id_vars=identifier_cols, value_vars=inds,
                                    var_name="Industry", value_name="FreshPop")
    caps = (panel_long.merge(fresh_long, on=["Region","Size","Industry"], how="outer")
                      .fillna({"PanelPop":0,"FreshPop":0}))

    cap_map = {(r,s,i):(pp,fp) for r,s,i,pp,fp in caps.to_records(index=False)}

    out = df_long_sol.copy()
    out["PanelAllocated"] = 0
    out["FreshAllocated"] = 0

    for i, row in out.iterrows():
        key = (row["Region"], row["Size"], row["Industry"])
        panel_cap, fresh_cap = cap_map.get(key, (0,0))
        x_i = int(row.get("OptimizedSample", 0))

        # greedy but SAFE split
        p_i = min(panel_cap, x_i)
        f_i = min(fresh_cap, x_i - p_i)
        out.at[i, "PanelAllocated"] = p_i
        out.at[i, "FreshAllocated"] = f_i

        # if (p_i + f_i) < x_i, we silently leave the remainder unassigned.
        # The solver-based path avoids this by construction.
    return out


###############################################################################
# 5) MAIN APP (SAME STRUCTURE, ADDED SAVE/LOAD SESSION)
###############################################################################

st.set_page_config(
    page_title="Survey Design",
    layout="wide",                      # <<< key
    initial_sidebar_state="expanded"
)
st.markdown("""
<style>
/* widen the main block and trim horizontal padding */
[data-testid="stAppViewContainer"] > .main {
  max-width: 100%;
  padding-left: 1rem;
  padding-right: 1rem;
}

/* also let the inner block container span wide */
.block-container {
  max-width: 100%;
  padding-top: 1rem;
  padding-bottom: 1rem;
}

/* optional: slightly tighter gap between columns */
[data-testid="stHorizontalBlock"] {
  gap: 0.75rem !important;
}
</style>
""", unsafe_allow_html=True)


def main():
    st.title("Survey Design")

    # --- Session Persistence in the Sidebar ---
    with st.sidebar.expander("Session Persistence", expanded=False):
        st.write("Save or load this session so others can review or resume later.")

        # Save Session
        if st.button("Save Session", key="btn_save_session"):
            session_id = str(uuid.uuid4())[:8]  # short random ID
            session_data = {
                "total_sample_1" : st.session_state.get("total_sample_1", 1000),
                "min_cell_size_1": st.session_state.get("min_cell_size_1", 4),
                "max_cell_size_1": st.session_state.get("max_cell_size_1", 40),
                "max_base_weight_1": st.session_state.get("max_base_weight_1", 600),
                "solver_choice_1": st.session_state.get("solver_choice_1", "SCIP"),
                "conversion_rate_1": st.session_state.get("conversion_rate_1", 0.3),
                "z_score_1": st.session_state.get("z_score_1", 1.644853627),
                "margin_of_error_1": st.session_state.get("margin_of_error_1", 0.075),
                "p_1": st.session_state.get("p_1", 0.5),

                "total_sample_2" : st.session_state.get("total_sample_2", 800),
                "min_cell_size_2": st.session_state.get("min_cell_size_2", 4),
                "max_cell_size_2": st.session_state.get("max_cell_size_2", 40),
                "max_base_weight_2": st.session_state.get("max_base_weight_2", 600),
                "solver_choice_2": st.session_state.get("solver_choice_2", "ECOS_BB"),
                "conversion_rate_2": st.session_state.get("conversion_rate_2", 0.3),
                "z_score_2": st.session_state.get("z_score_2", 1.644853627),
                "margin_of_error_2": st.session_state.get("margin_of_error_2", 0.075),
                "p_2": st.session_state.get("p_2", 0.5),

                "use_sum_universe": st.session_state.get("use_sum_universe", False),
            }
            save_session_to_file(session_id, session_data)
            st.success(f"Session saved! Session ID = {session_id}")
            # We can't use st.request.url, so here's a placeholder link
            st.markdown(f"Open with: https://surveydesigntest.streamlit.app/?session_id={session_id}")

        # Load Session
        load_id = st.text_input("Session ID to load", "")
        if st.button("Load Session", key="btn_load_session"):
            if not load_id.strip():
                st.warning("Please enter a session ID first.")
            else:
                loaded_data = load_session_from_file(load_id.strip())
                if loaded_data is None:
                    st.error(f"No saved session found for ID={load_id}")
                else:
                    st.success("Session loaded! Restoring parameters to session_state...")
                    # Put them into st.session_state
                    for k,v in loaded_data.items():
                        st.session_state[k] = v
                    # Because st.experimental_rerun is unavailable, we can't auto-refresh.
                    # So we just instruct user to refresh manually:
                    st.info("Session parameters have been restored. Please manually refresh the page to see changes.")


    st.write("""
    **Features**:
    1. Two sheets: 'panel' and 'fresh'.
    2. Checkbox: if unchecked => Adjusted Universe = max(panel,fresh), if checked => sum(panel+fresh).
    3. Check single-constraint feasibility. If there's a direct conflict, show the Overall, Cell, and Dimension conflicts as tables.
    4. If still we fail at the solver => run a slack-based diagnostic to see combined conflict.
    5. Two scenarios with separate parameters, plus difference.
    6. Dimension minimums auto-calculated from sample-size formula (override in sidebar).
    7. HTML & Excel downloads with color-scale in base-weight columns.
    """)

    all_solvers_list= ["SCIP","ECOS_BB"]

    # --- SCENARIO 1 parameters ---
    st.sidebar.header("Parameters for Scenario 1")
    total_sample_1= st.sidebar.number_input("Total Sample", value=1000, key="total_sample_1")
    min_cell_size_1= st.sidebar.number_input("Min Cell Size", value=4, key="min_cell_size_1")
    max_cell_size_1= st.sidebar.number_input("Max Cell Size", value=40, key="max_cell_size_1")
    max_base_weight_1= st.sidebar.number_input("Max Base Weight", value=600, key="max_base_weight_1")
    solver_choice_1= st.sidebar.selectbox("Solver", all_solvers_list, index=0, key="solver_choice_1")
    conversion_rate_1= st.sidebar.number_input("Conversion Rate", value=0.3, step=0.01, key="conversion_rate_1")

    use_sum_universe = st.sidebar.checkbox("Use sum(panel,fresh) instead of max(panel,fresh)", value=False)

    st.sidebar.markdown("---")
    st.sidebar.markdown("**Sample Size Formula Inputs (Scenario 1)**")
    z_score_1= st.sidebar.number_input("Z-Score", value=1.644853627, format="%.9f", key="z_score_1")
    margin_of_error_1= st.sidebar.number_input("Margin of Error", value=0.075, format="%.3f", key="margin_of_error_1")
    p_1= st.sidebar.number_input("p (Population Proportion)", value=0.5, format="%.2f", key="p_1")

    st.sidebar.markdown("---")

    # --- SCENARIO 2 parameters ---
    st.sidebar.header("Parameters for Scenario 2")
    total_sample_2= st.sidebar.number_input("Total Sample", value=800, key="total_sample_2")
    min_cell_size_2= st.sidebar.number_input("Min Cell Size", value=4, key="min_cell_size_2")
    max_cell_size_2= st.sidebar.number_input("Max Cell Size", value=40, key="max_cell_size_2")
    max_base_weight_2= st.sidebar.number_input("Max Base Weight", value=600, key="max_base_weight_2")
    solver_choice_2= st.sidebar.selectbox("Solver ", all_solvers_list, index=1, key="solver_choice_2")
    conversion_rate_2= st.sidebar.number_input("Conversion Rate ", value=0.3, step=0.01, key="conversion_rate_2")

    st.sidebar.markdown("**Sample Size Formula Inputs (Scenario 2)**")
    z_score_2= st.sidebar.number_input("Z-Score ", value=1.644853627, format="%.9f", key="z_score_2")
    margin_of_error_2= st.sidebar.number_input("Margin of Error ", value=0.075, format="%.3f", key="margin_of_error_2")
    p_2= st.sidebar.number_input("p (Population Proportion) ", value=0.5, format="%.2f", key="p_2")

    st.sidebar.markdown("---")

    uploaded_file= st.file_uploader("Upload Excel with 'panel','fresh'", type=["xlsx"])

    # We'll store dimension_mins for scenario1 and scenario2:
    dimension_mins_1= {"Region":{}, "Size":{}, "Industry":{}}
    dimension_mins_2= {"Region":{}, "Size":{}, "Industry":{}}

    if uploaded_file is not None:
        # Derive base_filename for display
        base_filename = uploaded_file.name.rsplit('.', 1)[0]
        if "_" in base_filename:
            parts = base_filename.split("_", 1)
            display_title = f"{parts[0]} for {parts[1]}"
        else:
            display_title = base_filename
        st.title(display_title)

        try:
            df_panel_wide= pd.read_excel(uploaded_file, sheet_name="panel")
            df_fresh_wide= pd.read_excel(uploaded_file, sheet_name="fresh")
        except Exception as e:
            st.error(f"Error reading 'panel'/'fresh' => {e}")
            return

        try:
            df_frame_panel_wide = pd.read_excel(uploaded_file, sheet_name="frame_panel")
            df_frame_fresh_wide = pd.read_excel(uploaded_file, sheet_name="frame_fresh")
        except Exception as e:
            st.error(f"Error reading 'frame_panel'/'frame_fresh' => {e}")
            return

        
        # Show as expanders (dropdowns)
        with st.expander("Frame – Panel (frame_panel)"):
            st.data_editor(df_frame_panel_wide, use_container_width=True, key="frame_panel_table")
        
        with st.expander("Frame – Fresh (frame_fresh)"):
            st.data_editor(df_frame_fresh_wide, use_container_width=True, key="frame_fresh_table")
        
        # Optional: one quick dropdown viewer for any input table
        with st.expander("Quick Table Viewer"):
            table_choice = st.selectbox(
                "Pick a table to preview",
                ["panel", "fresh", "frame_panel", "frame_fresh"],
                index=0,
                key="table_choice"
            )
            _map = {
                "panel": df_panel_wide,
                "fresh": df_fresh_wide,
                "frame_panel": df_frame_panel_wide,
                "frame_fresh": df_frame_fresh_wide,
            }
            st.data_editor(_map[table_choice], use_container_width=True, key="table_choice_view")
        id_cols = ["Region","Size"]
        industry_cols = [c for c in df_panel_wide.columns if c not in id_cols]
        
        missing_fp = set(industry_cols) - set(df_frame_panel_wide.columns)
        missing_ff = set(industry_cols) - set(df_frame_fresh_wide.columns)
        if missing_fp:
            st.warning(f"'frame_panel' missing industries: {sorted(missing_fp)}")
        if missing_ff:
            st.warning(f"'frame_fresh' missing industries: {sorted(missing_ff)}")
        
        with st.expander("Original input tables (edit as needed)"):
            show_side_by_side(
                df_panel_wide, "Panel (input)",
                df_fresh_wide, "Fresh (input)",
                key_left="wide_panel_scenario",
                key_right="wide_fresh_scenario",
                editable=True
            )

#        with st.expander("Original Panel Table"):
#            st.data_editor(df_panel_wide, use_container_width=True, key="wide_panel_scenario")
 #       with st.expander("Original Fresh Table"):
  #          st.data_editor(df_fresh_wide, use_container_width=True, key="wide_fresh_scenario")

        # build Adjusted Universe
        df_adjusted= df_panel_wide.copy()
        id_cols= ["Region","Size"]
        industry_cols= [c for c in df_panel_wide.columns if c not in id_cols]
        for c in industry_cols:
            if c in df_fresh_wide.columns:
                if use_sum_universe:
                    df_adjusted[c] = df_panel_wide[c].fillna(0) + df_fresh_wide[c].fillna(0)
                else:
                    df_adjusted[c] = np.maximum(df_panel_wide[c].fillna(0), df_fresh_wide[c].fillna(0))

        st.subheader("Adjusted Universe Table")
        st.data_editor(df_adjusted, use_container_width=True, key="adjusted_universe_scenario")

        # dimension sets from df_adjusted
        all_regions= df_adjusted["Region"].dropna().unique()
        all_sizes= df_adjusted["Size"].dropna().unique()
        all_inds= industry_cols

        # sample-size formulas for scenario 1
        n_infinity_1 = compute_n_infinity(z_score_1, margin_of_error_1, p_1)
        def sum_pop_in_dim_1(df, dim_type, val):
            subset= df[df[dim_type]== val]
            tot=0
            for cc in all_inds:
                tot+= subset[cc].fillna(0).sum()
            return tot

        # sample-size formulas for scenario 2
        n_infinity_2 = compute_n_infinity(z_score_2, margin_of_error_2, p_2)
        def sum_pop_in_dim_2(df, dim_type, val):
            subset= df[df[dim_type]== val]
            tot=0
            for cc in all_inds:
                tot+= subset[cc].fillna(0).sum()
            return tot

        with st.sidebar.expander("Dimension Minimum", expanded=True):
            st.write("### Scenario 1")
            for r in all_regions:
                pop_ = sum_pop_in_dim_1(df_adjusted,"Region", r)
                defMin= compute_fpc_min(pop_, n_infinity_1)
                user_val= st.number_input(f"**Region**={r} (S1)", 
                                          min_value=0, value=int(round(defMin)), step=1, key=f"S1_reg_{r}")
                dimension_mins_1["Region"][r]= user_val
            for sz in all_sizes:
                pop_ = sum_pop_in_dim_1(df_adjusted,"Size", sz)
                defMin= compute_fpc_min(pop_, n_infinity_1)
                user_val= st.number_input(f"**Size**={sz} (S1)",
                                          min_value=0, value=int(round(defMin)), step=1, key=f"S1_size_{sz}")
                dimension_mins_1["Size"][sz]= user_val
            for ind_ in all_inds:
                pop_= df_adjusted[ind_].fillna(0).sum()
                defMin= compute_fpc_min(pop_, n_infinity_1)
                user_val= st.number_input(f"**Sector**={ind_} (S1)",
                                          min_value=0, value=int(round(defMin)), step=1, key=f"S1_ind_{ind_}")
                dimension_mins_1["Industry"][ind_]= user_val

            st.write("### Scenario 2")
            for r in all_regions:
                pop_ = sum_pop_in_dim_2(df_adjusted,"Region", r)
                defMin= compute_fpc_min(pop_, n_infinity_2)
                user_val= st.number_input(f"**Region**={r} (S2)", 
                                          min_value=0, value=int(round(defMin)), step=1, key=f"S2_reg_{r}")
                dimension_mins_2["Region"][r]= user_val
            for sz in all_sizes:
                pop_ = sum_pop_in_dim_2(df_adjusted,"Size", sz)
                defMin= compute_fpc_min(pop_, n_infinity_2)
                user_val= st.number_input(f"**Size**={sz} (S2)",
                                          min_value=0, value=int(round(defMin)), step=1, key=f"S2_size_{sz}")
                dimension_mins_2["Size"][sz]= user_val
            for ind_ in all_inds:
                pop_= df_adjusted[ind_].fillna(0).sum()
                defMin= compute_fpc_min(pop_, n_infinity_2)
                user_val= st.number_input(f"**Sector**={ind_} (S2)",
                                          min_value=0, value=int(round(defMin)), step=1, key=f"S2_ind_{ind_}")
                dimension_mins_2["Industry"][ind_]= user_val

        # ---------- RUN BOTH SCENARIOS -----------
        if st.button("Run Optimization for Both Scenarios"):
            # define a quick function to run scenario
            def run_scenario(label, 
                             total_sample, min_cell, max_cell, max_bw, solver, conv_rate,
                             dimension_mins):
                """
                1) run_optimization
                2) if success => allocate_panel_fresh => produce pivot tables
                else => handle conflicts or do slack diag
                """
                scenario_result = {}
                try:
                    df_long_final, df_cell_conf, df_dim_conf, solver_info = run_optimization(
                        df_wide=df_adjusted,
                        total_sample=total_sample,
                        min_cell_size=min_cell,
                        max_cell_size=max_cell,
                        max_base_weight=max_bw,
                        solver_choice=solver,
                        dimension_mins=dimension_mins,
                        conversion_rate=conv_rate,
                        df_frame_panel_wide=df_frame_panel_wide,
                        df_frame_fresh_wide=df_frame_fresh_wide,
                    )
                
                    if df_long_final is None:
                        # single-constraint conflict
                        scenario_result["success"] = False
                        scenario_result["cell_conflicts"] = df_cell_conf
                        scenario_result["dim_conflicts"]  = df_dim_conf
                        scenario_result["overall_conflicts"] = solver_info
                        scenario_result["solver_info"] = None
                    else:
                        scenario_result["success"] = True
                        scenario_result["solver_info"] = solver_info
                        # allocate
#                        df_alloc = allocate_panel_fresh(df_long_final, df_panel_wide, df_fresh_wide)
 #                       scenario_result["df_alloc"] = df_alloc
                        # allocations come from the solver (feasible by construction)
                        df_alloc = df_long_final

                        # region totals, size totals
                        region_totals = (
                            df_alloc
                            .groupby("Region")[["PanelAllocated","FreshAllocated"]]
                            .sum()
                            .assign(SampleTotal=lambda d: d["PanelAllocated"] + d["FreshAllocated"])
                            .reset_index()
                        )
                        size_totals = (
                            df_alloc
                            .groupby("Size")[["PanelAllocated","FreshAllocated"]]
                            .sum()
                            .assign(SampleTotal=lambda d: d["PanelAllocated"] + d["FreshAllocated"])
                            .reset_index()
                        )

                        pivot_panel = pd.pivot_table(
                            df_alloc,
                            index=["Region","Size"],
                            columns="Industry",
                            values="PanelAllocated",
                            aggfunc='sum',
                            fill_value=0,
                            margins=True,
                            margins_name="GrandTotal",
                            sort=False
                        ).reset_index()

                        pivot_fresh = pd.pivot_table(
                            df_alloc,
                            index=["Region","Size"],
                            columns="Industry",
                            values="FreshAllocated",
                            aggfunc='sum',
                            fill_value=0,
                            margins=True,
                            margins_name="GrandTotal",
                            sort=False
                        ).reset_index()

                        df_combined = create_combined_table_with_totals(df_alloc)
                        df_combined = df_combined.reset_index()

                        # reorder columns for df_combined
                        id_cols = [c for c in ["Region","Size"] if c in df_combined.columns]
                        gt_sample = [c for c in df_combined.columns if c=="GrandTotal_Sample"]
                        gt_bw     = [c for c in df_combined.columns if c=="GrandTotal_BaseWeight"]
                        sample_cols = [c for c in df_combined.columns if c.endswith("_Sample") and c not in gt_sample]
                        bw_cols     = [c for c in df_combined.columns if c.endswith("_BaseWeight") and c not in gt_bw]
                        new_order = id_cols + gt_sample + sample_cols + bw_cols + gt_bw
                        df_combined = df_combined[new_order]

                        # round baseweight columns
                        subset_bw_cols = [c for c in df_combined.columns if c.endswith("_BaseWeight")]
                        df_combined[subset_bw_cols] = df_combined[subset_bw_cols].round(1)

                        # color-scale
                        norm_bw_cols = [c for c in subset_bw_cols if c!="GrandTotal_BaseWeight"]
                        if len(df_combined)>1 and len(norm_bw_cols)>0:
                            norm_df = df_combined.iloc[:-1]
                            global_min = norm_df[norm_bw_cols].min().min()
                            global_max = norm_df[norm_bw_cols].max().max()
                            global_mid = np.percentile(norm_df[norm_bw_cols].stack(), 50)
                        else:
                            global_min=0
                            global_mid=0
                            global_max=0
                        custom_cmap = LinearSegmentedColormap.from_list("custom", ["#00FF00","#FFFF00","#FF0000"])
                        norm_obj = TwoSlopeNorm(vmin=global_min,vcenter=global_mid,vmax=global_max)

                        def baseweight_color(v):
                            val = norm_obj(v)
                            color = mcolors.to_hex(custom_cmap(val))
                            return f"background-color: {color}"
                        def style_bwcol(series):
                            n=len(series)
                            return [baseweight_color(x) if i<n-1 else "" for i,x in enumerate(series)]

                        stcol = df_combined.style.apply(style_bwcol, subset=norm_bw_cols)

                        # population pivot
                        pivot_pop = pd.pivot_table(
                            df_alloc,
                            index=["Region","Size"],
                            columns="Industry",
                            values="Population",
                            aggfunc='sum',
                            fill_value=0,
                            margins=True,
                            margins_name="GrandTotal",
                            sort=False
                        ).reset_index()

                        # proportional pivot
                        pivot_propsample= pd.pivot_table(
                            df_alloc,
                            index=["Region","Size"],
                            columns="Industry",
                            values="PropSample",
                            aggfunc='mean',
                            fill_value=0,
                            margins=False,
                            sort=False
                        ).reset_index()
                        num_cols = pivot_propsample.select_dtypes(include="number").columns
                        pivot_propsample[num_cols] = (
                            pivot_propsample[num_cols].round(0).fillna(0).astype(int)
                        )

                        scenario_result["region_totals"] = region_totals
                        scenario_result["size_totals"]   = size_totals
                        scenario_result["pivot_fresh"]   = pivot_fresh
                        scenario_result["pivot_panel"]   = pivot_panel
                        scenario_result["df_combined"]   = df_combined
                        scenario_result["df_combined_style"] = stcol
                        scenario_result["pivot_pop"]     = pivot_pop
                        scenario_result["pivot_propsample"] = pivot_propsample

                except ValueError as e:
                    if "No solver found a feasible solution" in str(e):
                        # do slack diag
                        df_long = df_adjusted.melt(
                            id_vars=["Region","Size"],
                            var_name="Industry",
                            value_name="Population"
                        ).fillna(0)
                        diag_sol, diag_usage, diag_status = diagnose_infeasibility_slacks(
                            df_long,
                            total_sample,
                            min_cell,
                            max_cell,
                            max_bw,
                            dimension_mins,
                            conv_rate
                        )

                        df_long_diag = (df_adjusted.melt(id_vars=["Region","Size"], var_name="Industry", value_name="Population")
                                        .merge(df_panel_wide.melt(id_vars=["Region","Size"], value_vars=industry_cols,
                                                                  var_name="Industry", value_name="PanelPop"),
                                               on=["Region","Size","Industry"], how="left")
                                        .merge(df_fresh_wide.melt(id_vars=["Region","Size"], value_vars=industry_cols,
                                                                  var_name="Industry", value_name="FreshPop"),
                                               on=["Region","Size","Industry"], how="left")
                                        .fillna({"PanelPop":0,"FreshPop":0}))
                        diag_sol, diag_usage, diag_status = diagnose_infeasibility_slacks(
                            df_long_diag, total_sample, min_cell, max_cell, max_bw, dimension_mins, conv_rate
                        )

                        scenario_result["success"]=False
                        scenario_result["error_msg"] = str(e)
                        scenario_result["diag_sol"] = diag_sol
                        scenario_result["diag_usage"] = diag_usage
                        scenario_result["diag_status"] = diag_status
                    else:
                        scenario_result["success"]=False
                        scenario_result["error_msg"] = str(e)
                except Exception as e2:
                    scenario_result["success"] = False
                    scenario_result["error_msg"] = f"Solver Error: {e2}"
                return scenario_result

            # run scenario 1
            scenario1_result = run_scenario(
                label="Scenario1_",
                total_sample=total_sample_1,
                min_cell=min_cell_size_1,
                max_cell=max_cell_size_1,
                max_bw=max_base_weight_1,
                solver=solver_choice_1,
                conv_rate=conversion_rate_1,
                dimension_mins=dimension_mins_1
            )
            params_df = pd.DataFrame([
            ("Total Sample",        total_sample_1),
            ("Min Cell Size",       min_cell_size_1),
            ("Max Cell Size",       max_cell_size_1),
            ("Max Base Weight",     max_base_weight_1),
            ("Solver Choice",       solver_choice_1),
            ("Conversion Rate",     conversion_rate_1),
            ("Z-Score",             z_score_1),
            ("Margin of Error",     margin_of_error_1),
            ("p",                   p_1),
            ], columns=["Parameter","Value"])
            # Then we can build dataframes for each dimension:
            region_min_df = pd.DataFrame(list(dimension_mins_1["Region"].items()),
                                         columns=["Region", "MinNeeded"])
            size_min_df = pd.DataFrame(list(dimension_mins_1["Size"].items()),
                                       columns=["Size", "MinNeeded"])
            industry_min_df = pd.DataFrame(list(dimension_mins_1["Industry"].items()),
                                           columns=["Industry", "MinNeeded"])

            # run scenario 2
            scenario2_result = run_scenario(
                label="Scenario2_",
                total_sample=total_sample_2,
                min_cell=min_cell_size_2,
                max_cell=max_cell_size_2,
                max_bw=max_base_weight_2,
                solver=solver_choice_2,
                conv_rate=conversion_rate_2,
                dimension_mins=dimension_mins_2
            )

            params_df2 = pd.DataFrame([
                ("Total Sample",        total_sample_2),
                ("Min Cell Size",       min_cell_size_2),
                ("Max Cell Size",       max_cell_size_2),
                ("Max Base Weight",     max_base_weight_2),
                ("Solver Choice",       solver_choice_2),
                ("Conversion Rate",     conversion_rate_2),
                ("Z-Score",             z_score_2),
                ("Margin of Error",     margin_of_error_2),
                ("p",                   p_2),
            ], columns=["Parameter","Value"])
            
            region_min_df2 = pd.DataFrame(list(dimension_mins_2["Region"].items()),
                                          columns=["Region", "MinNeeded"])
            size_min_df2 = pd.DataFrame(list(dimension_mins_2["Size"].items()),
                                        columns=["Size", "MinNeeded"])
            industry_min_df2 = pd.DataFrame(list(dimension_mins_2["Industry"].items()),
                                            columns=["Industry", "MinNeeded"])

            # show scenario 1
            st.header("Scenario 1 Results")
            if not scenario1_result["success"]:
                # conflict or solver error
                if "cell_conflicts" in scenario1_result:
                    st.error("Single-constraint conflict(s) for Scenario 1.")
                    df_overall_1 = scenario1_result["overall_conflicts"]
                    df_dim_conf_1= scenario1_result["dim_conflicts"]
                    df_cell_conf_1= scenario1_result["cell_conflicts"]
                    if df_overall_1 is not None and not df_overall_1.empty:
                        st.subheader("Overall Conflicts (Scenario 1)")
                        st.data_editor(df_overall_1, use_container_width=True, key="s1_overall_conf")
                    if df_dim_conf_1 is not None and not df_dim_conf_1.empty:
                        st.subheader("Dimension Conflicts (Scenario 1)")
                        st.data_editor(df_dim_conf_1, use_container_width=True, key="s1_dim_conf")
                    if df_cell_conf_1 is not None and not df_cell_conf_1.empty:
                        st.subheader("Cell Conflicts (Scenario 1)")
                        st.data_editor(df_cell_conf_1, use_container_width=True, key="s1_cell_conf")
                elif "error_msg" in scenario1_result:
                    st.error(scenario1_result["error_msg"])
                    if "diag_sol" in scenario1_result:
                        st.warning("Slack-based diagnostic for Scenario 1:")
                        diag_sol1 = scenario1_result["diag_sol"]
                        diag_usage1 = scenario1_result.get("diag_usage", pd.DataFrame())
                        diag_status1= scenario1_result.get("diag_status", "")
                        if diag_sol1 is not None:
                            st.dataframe(diag_sol1, key="s1_diag_sol")
                        if diag_usage1 is not None and not diag_usage1.empty:
                            st.dataframe(diag_usage1, key="s1_diag_usage")
                        if diag_status1:
                            st.write(f"Diagnostic solver status: {diag_status1}")
            else:
                # success
                solver1_info = scenario1_result["solver_info"]
                st.success(f"Solved with solver: {solver1_info}")

                show_side_by_side(
                    scenario1_result["pivot_panel"], "Panel (Scenario 1)",
                    scenario1_result["pivot_fresh"], "Fresh (Scenario 1)",
                    key_left="s1_panel", key_right="s1_fresh", editable=False
                )

                
                st.subheader("Allocated Sample & Base Weights (Scenario 1)")
                st.dataframe(scenario1_result["df_combined"], key="s1_combined")

                st.subheader("Region-wise Sample Totals (Scenario 1)")
                st.data_editor(scenario1_result["region_totals"], use_container_width=True, key="s1_region_totals")

                st.subheader("Size-wise Sample Totals (Scenario 1)")
                st.data_editor(scenario1_result["size_totals"], use_container_width=True, key="s1_size_totals")

                # Region totals side-by-side (panel vs fresh)
                reg_panel = scenario1_result["region_totals"][["Region","PanelAllocated"]].rename(columns={"PanelAllocated":"Panel"})
                reg_fresh = scenario1_result["region_totals"][["Region","FreshAllocated"]].rename(columns={"FreshAllocated":"Fresh"})
                show_side_by_side(reg_panel, "Region totals – Panel (S1)", reg_fresh, "Region totals – Fresh (S1)",
                                  key_left="s1_reg_panel", key_right="s1_reg_fresh", editable=False)
                
                # Size totals side-by-side (panel vs fresh)
                sz_panel = scenario1_result["size_totals"][["Size","PanelAllocated"]].rename(columns={"PanelAllocated":"Panel"})
                sz_fresh = scenario1_result["size_totals"][["Size","FreshAllocated"]].rename(columns={"FreshAllocated":"Fresh"})
                show_side_by_side(sz_panel, "Size totals – Panel (S1)", sz_fresh, "Size totals – Fresh (S1)",
                                  key_left="s1_size_panel", key_right="s1_size_fresh", editable=False)



            # show scenario 2
            st.header("Scenario 2 Results")
            if not scenario2_result["success"]:
                # conflict or solver error
                if "cell_conflicts" in scenario2_result:
                    st.error("Single-constraint conflict(s) for Scenario 2.")
                    df_overall_2 = scenario2_result["overall_conflicts"]
                    df_dim_conf_2= scenario2_result["dim_conflicts"]
                    df_cell_conf_2= scenario2_result["cell_conflicts"]
                    if df_overall_2 is not None and not df_overall_2.empty:
                        st.subheader("Overall Conflicts (Scenario 2)")
                        st.data_editor(df_overall_2, use_container_width=True, key="s2_overall_conf")
                    if df_dim_conf_2 is not None and not df_dim_conf_2.empty:
                        st.subheader("Dimension Conflicts (Scenario 2)")
                        st.data_editor(df_dim_conf_2, use_container_width=True, key="s2_dim_conf")
                    if df_cell_conf_2 is not None and not df_cell_conf_2.empty:
                        st.subheader("Cell Conflicts (Scenario 2)")
                        st.data_editor(df_cell_conf_2, use_container_width=True, key="s2_cell_conf")
                elif "error_msg" in scenario2_result:
                    st.error(scenario2_result["error_msg"])
                    if "diag_sol" in scenario2_result:
                        st.warning("Slack-based diagnostic for Scenario 2:")
                        diag_sol2 = scenario2_result["diag_sol"]
                        diag_usage2 = scenario2_result.get("diag_usage", pd.DataFrame())
                        diag_status2= scenario2_result.get("diag_status", "")
                        if diag_sol2 is not None:
                            st.dataframe(diag_sol2, key="s2_diag_sol")
                        if not diag_usage2.empty:
                            st.dataframe(diag_usage2, key="s2_diag_usage")
                        if diag_status2:
                            st.write(f"Diagnostic solver status: {diag_status2}")
            else:
                # success
                solver2_info = scenario2_result["solver_info"]
                st.success(f"Solved with solver: {solver2_info}")
               
                show_side_by_side(
                    scenario2_result["pivot_panel"], "Panel (Scenario 2)",
                    scenario2_result["pivot_fresh"], "Fresh (Scenario 2)",
                    key_left="s2_panel", key_right="s2_fresh", editable=False
                )

                
                st.subheader("Allocated Sample & Base Weights (Scenario 2)")
                st.dataframe(scenario2_result["df_combined_style"], key="s2_combined")

                st.subheader("Region-wise Sample Totals (Scenario 2)")
                st.data_editor(scenario2_result["region_totals"], use_container_width=True, key="s2_region_totals")

                st.subheader("Size-wise Sample Totals (Scenario 2)")
                st.data_editor(scenario2_result["size_totals"], use_container_width=True, key="s2_size_totals")

                reg_panel2 = scenario2_result["region_totals"][["Region","PanelAllocated"]].rename(columns={"PanelAllocated":"Panel"})
                reg_fresh2 = scenario2_result["region_totals"][["Region","FreshAllocated"]].rename(columns={"FreshAllocated":"Fresh"})
                show_side_by_side(reg_panel2, "Region totals – Panel (S2)", reg_fresh2, "Region totals – Fresh (S2)",
                                  key_left="s2_reg_panel", key_right="s2_reg_fresh", editable=False)
                
                sz_panel2 = scenario2_result["size_totals"][["Size","PanelAllocated"]].rename(columns={"PanelAllocated":"Panel"})
                sz_fresh2 = scenario2_result["size_totals"][["Size","FreshAllocated"]].rename(columns={"FreshAllocated":"Fresh"})
                show_side_by_side(sz_panel2, "Size totals – Panel (S2)", sz_fresh2, "Size totals – Fresh (S2)",
                                  key_left="s2_size_panel", key_right="s2_size_fresh", editable=False)


            

            # -------------- COMPARISON --------------
            st.header("Comparison: Scenario 1 minus Scenario 2")
            if scenario1_result.get("success") and scenario2_result.get("success"):
                dfc1 = scenario1_result["df_combined"].copy().reset_index(drop=True)
                dfc2 = scenario2_result["df_combined"].copy().reset_index(drop=True)

                dfc1 = scenario1_result["df_combined"].copy().reset_index()
                dfc2 = scenario2_result["df_combined"].copy().reset_index()

                common_cols = [c for c in dfc1.columns if c in dfc2.columns]
                dfc1 = dfc1[common_cols]
                dfc2 = dfc2[common_cols]
                numeric_cols = [c for c in common_cols if pd.api.types.is_numeric_dtype(dfc1[c])]
                df_diff = dfc1.copy()
                for col in numeric_cols:
                    df_diff[col] = dfc1[col] - dfc2[col]

                def color_diff(val):
                    if pd.isna(val):
                        return ""
                    if val>0:
                        return "background-color:#b3ffb3"
                    elif val<0:
                        return "background-color:#ffb3b3"
                    else:
                        return ""
                df_diff_style = df_diff.style.applymap(color_diff, subset=numeric_cols)
                st.dataframe(df_diff_style, key="comparison_df_diff")

            # ---------- HTML & Excel downloads ----------
            # We'll collect scenario sections for HTML
            scenario_sections = []
            # scenario 1
            if scenario1_result.get("success"):
                scenario_sections.append(("Scenario 1: Allocated Sample & Base Weights",
                                          scenario1_result["df_combined_style"]))
            # scenario 2
            if scenario2_result.get("success"):
                scenario_sections.append(("Scenario 2: Allocated Sample & Base Weights",
                                          scenario2_result["df_combined_style"]))
            # difference
            if scenario1_result.get("success") and scenario2_result.get("success"):
                scenario_sections.append(("Scenario 1 vs. 2 Difference", df_diff_style))

            if scenario_sections:
                html_bytes = dfs_to_html(scenario_sections, page_title=display_title)
                html_fname = f"{base_filename}_comparison_snapshot.html"
                st.download_button(
                    label="🌐 Download comparison as HTML",
                    data=html_bytes,
                    file_name=html_fname,
                    mime="text/html"
                )

            ###############################################################################
            # FINAL EXCEL EXPORT CODE: Region & Size columns in baseweight & difference
            # with color scale for scenario baseweights, no proportional-sample sheets
            ###############################################################################
            
            # 1) List of industry columns in the original input order
            industries_in_input = [
                c for c in df_panel_wide.columns
                if c not in ("Region","Size")
            ]

            
            def reorder_and_color_scenario(df_in, sheet_name, writer):
                """
                For a scenario's allocated sample df_in, we reorder columns:
                  Region,Size (if present),
                  then for each industry in df_panel_wide => Ind_Sample, Ind_BaseWeight,
                  then GrandTotal_Sample, GrandTotal_BaseWeight if present.
                Applies color scale to base-weight columns (excluding GrandTotal).
                """
                df_out = df_in.reset_index(drop=True).copy()
            
                # Step A: Build column order
                col_order = []
                # region & size
                for c in ("Region","Size"):
                    if c in df_out.columns:
                        col_order.append(c)
            
                # sample & baseweight columns in industry order
                sample_cols = []
                bw_cols     = []
                for ind_ in industries_in_input:
                    s_col = f"{ind_}_Sample"
                    if s_col in df_out.columns:
                        sample_cols.append(s_col)
                    b_col = f"{ind_}_BaseWeight"
                    if b_col in df_out.columns:
                        bw_cols.append(b_col)
            
                # grand-totals last
                extras = []
                for c_ in ["GrandTotal_Sample","GrandTotal_BaseWeight"]:
                    if c_ in df_out.columns:
                        extras.append(c_)
            
                col_order.extend(sample_cols)
                col_order.extend(bw_cols)
                col_order.extend(extras)
            
                # intersection to handle missing columns
                col_order = [c for c in col_order if c in df_out.columns]
            
                # reorder
                df_out = df_out[col_order]
            
                # Step B: Write to Excel
                df_out.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]
            
                # Step C: Color-scale for base-weight columns
                from openpyxl.utils import get_column_letter
                from openpyxl.formatting.rule import ColorScaleRule
            
                real_bw_cols = [c for c in bw_cols if c in df_out.columns and c!="GrandTotal_BaseWeight"]
                if len(df_out) > 1 and real_bw_cols:
                    df_no_total = df_out.iloc[:-1]  # ignoring last row if it's grand total
                    global_min = df_no_total[real_bw_cols].min().min()
                    global_max = df_no_total[real_bw_cols].max().max()
                    global_mid = np.percentile(df_no_total[real_bw_cols].stack(), 50)
                else:
                    global_min=0; global_mid=0; global_max=0
            
                def make_rule():
                    return ColorScaleRule(
                        start_type="num", start_value=global_min, start_color="00FF00",
                        mid_type="num",   mid_value=global_mid,  mid_color="FFFF00",
                        end_type="num",   end_value=global_max,  end_color="FF0000",
                    )
            
                n_rows = df_out.shape[0]
                for col_name in real_bw_cols:
                    col_idx = df_out.columns.get_loc(col_name) + 1  # 1-based for Excel
                    excel_col = get_column_letter(col_idx)
                    rng = f"{excel_col}2:{excel_col}{n_rows}"
                    ws.conditional_formatting.add(rng, make_rule())
                    # numeric formatting
                    for cell in ws[f"{excel_col}2":f"{excel_col}{n_rows}"]:
                        cell[0].number_format = "0.0"
            
            
            def reorder_diff_sheet(df_in):
                """
                For the difference sheet, includes Region/Size columns,
                then for each industry => Ind_Sample, Ind_BaseWeight,
                plus GrandTotals last if present.
                No color scale is applied to difference.
                """
                df_out = df_in.reset_index(drop=True).copy()
            
                # region & size first
                col_order = []
                for c in ("Region","Size"):
                    if c in df_out.columns:
                        col_order.append(c)
            
                sample_cols = []
                bw_cols     = []
                for ind_ in industries_in_input:
                    s_col = f"{ind_}_Sample"
                    if s_col in df_out.columns:
                        sample_cols.append(s_col)
                    b_col = f"{ind_}_BaseWeight"
                    if b_col in df_out.columns:
                        bw_cols.append(b_col)
            
                extras = []
                for c_ in ["GrandTotal_Sample","GrandTotal_BaseWeight"]:
                    if c_ in df_out.columns:
                        extras.append(c_)
            
                col_order.extend(sample_cols)
                col_order.extend(bw_cols)
                col_order.extend(extras)
            
                col_order = [c for c in col_order if c in df_out.columns]
                return df_out[col_order]
        
            # grab the exact Region/Size order from df_adjusted
            region_order = df_adjusted["Region"].dropna().drop_duplicates().tolist()
            size_order   = df_adjusted["Size"].dropna().drop_duplicates().tolist()
            
            excel_out = io.BytesIO()
            with pd.ExcelWriter(excel_out, engine="openpyxl") as writer:
            
                # 1) Adjusted Universe
                df_adjusted.to_excel(writer, sheet_name="Adjusted_Universe", index=False)
            
                # 2) Parameters & Cell Minimums side-by-side
                params_comp = pd.DataFrame({
                    "Parameter": params_df["Parameter"],
                    "S1_Value":  params_df["Value"],
                    "S2_Value":  params_df2["Value"],
                })
            
                # merge and then reindex to match df_adjusted order
                region_mins = pd.merge(
                    region_min_df.rename(columns={"MinNeeded":"S1_MinNeeded"}),
                    region_min_df2.rename(columns={"MinNeeded":"S2_MinNeeded"}),
                    on="Region", how="outer"
                ).set_index("Region").reindex(region_order).reset_index()
            
                size_mins = pd.merge(
                    size_min_df.rename(columns={"MinNeeded":"S1_MinNeeded"}),
                    size_min_df2.rename(columns={"MinNeeded":"S2_MinNeeded"}),
                    on="Size", how="outer"
                ).set_index("Size").reindex(size_order).reset_index()
            
                # industry order left as-is
                industry_mins = pd.merge(
                    industry_min_df.rename(columns={"MinNeeded":"S1_MinNeeded"}),
                    industry_min_df2.rename(columns={"MinNeeded":"S2_MinNeeded"}),
                    on="Industry", how="outer"
                )
            
                sheet_pm = "ParametersAndMins"
                params_comp.to_excel(writer, sheet_name=sheet_pm, index=False, startrow=0)
                r = params_comp.shape[0] + 2
                region_mins.to_excel(writer, sheet_name=sheet_pm, index=False, startrow=r)
                r += region_mins.shape[0] + 2
                size_mins.to_excel(writer, sheet_name=sheet_pm, index=False, startrow=r)
                r += size_mins.shape[0] + 2
                industry_mins.to_excel(writer, sheet_name=sheet_pm, index=False, startrow=r)

                # ---- Side-by-side sheets with mins (Scenario 1) ----
                if scenario1_result.get("success"):
                    # S1-specific mins (already built above): region_min_df, size_min_df, industry_min_df
                    write_panel_fresh_and_mins_sheet(
                        writer=writer,
                        sheet_name="S1_PanelFresh+Mins",
                        panel_df=scenario1_result["pivot_panel"],
                        fresh_df=scenario1_result["pivot_fresh"],
                        region_mins_df=region_min_df.rename(columns={"MinNeeded": "MinNeeded"}),
                        size_mins_df=size_min_df.rename(columns={"MinNeeded": "MinNeeded"}),
                        industry_mins_df=industry_min_df.rename(columns={"MinNeeded": "MinNeeded"}),
                        title_left="Panel (Scenario 1)",
                        title_right="Fresh (Scenario 1)",
                    )
                
                # ---- Side-by-side sheets with mins (Scenario 2) ----
                if scenario2_result.get("success"):
                    write_panel_fresh_and_mins_sheet(
                        writer=writer,
                        sheet_name="S2_PanelFresh+Mins",
                        panel_df=scenario2_result["pivot_panel"],
                        fresh_df=scenario2_result["pivot_fresh"],
                        region_mins_df=region_min_df2.rename(columns={"MinNeeded": "MinNeeded"}),
                        size_mins_df=size_min_df2.rename(columns={"MinNeeded": "MinNeeded"}),
                        industry_mins_df=industry_min_df2.rename(columns={"MinNeeded": "MinNeeded"}),
                        title_left="Panel (Scenario 2)",
                        title_right="Fresh (Scenario 2)",
                    )

            
                # 3) All Samples & BaseWeights (S1 & S2 stacked)
                def reorder_samples(df):
                    df_out = df.reset_index().copy()
                    base_cols = [c for c in ("Region","Size") if c in df_out.columns]
                    sample_cols, bw_cols = [], []
                    for ind in industries_in_input:
                        s, b = f"{ind}_Sample", f"{ind}_BaseWeight"
                        if s in df_out.columns: sample_cols.append(s)
                        if b in df_out.columns: bw_cols.append(b)
                    extras = [c for c in ("GrandTotal_Sample","GrandTotal_BaseWeight")
                              if c in df_out.columns]
                    return df_out[base_cols + sample_cols + bw_cols + extras]
            
                s1_sb = reorder_samples(scenario1_result["df_combined"])
                s2_sb = reorder_samples(scenario2_result["df_combined"])
            
                sheet_sb = "All_Samples_and_BaseWeights"
                s1_sb.to_excel(writer, sheet_name=sheet_sb, startrow=0, index=False)
                start_sb = s1_sb.shape[0] + 2
                s2_sb.to_excel(writer, sheet_name=sheet_sb, startrow=start_sb, index=False)
            
                ws_sb = writer.sheets[sheet_sb]
                def apply_color_scale(df_block, row_offset):
                    real_bw = [c for c in df_block.columns
                               if c.endswith("_BaseWeight") and c!="GrandTotal_BaseWeight"]
                    if len(df_block)>1 and real_bw:
                        sub = df_block.iloc[:-1]
                        vmin = sub[real_bw].min().min()
                        vmax = sub[real_bw].max().max()
                        vmid = np.percentile(sub[real_bw].stack(),50)
                        rule = ColorScaleRule(
                            start_type="num", start_value=vmin, start_color="00FF00",
                            mid_type="num",   mid_value=vmid,  mid_color="FFFF00",
                            end_type="num",   end_value=vmax,  end_color="FF0000",
                        )
                        n = df_block.shape[0]
                        top = row_offset + 2
                        bottom = row_offset + n
                        for col in real_bw:
                            idx = df_block.columns.get_loc(col)+1
                            letter = get_column_letter(idx)
                            ws_sb.conditional_formatting.add(f"{letter}{top}:{letter}{bottom}", rule)
                            for cell in ws_sb[f"{letter}{top}":f"{letter}{bottom}"]:
                                cell[0].number_format = "0.0"
            
                apply_color_scale(s1_sb, row_offset=0)
                apply_color_scale(s2_sb, row_offset=start_sb)
            
                # 4) Panel Allocations (S1 & S2 stacked)
                def reorder_panel(df):
                    cols = [c for c in ("Region","Size") if c in df.columns]
                    inds = [c for c in industries_in_input if c in df.columns]
                    extras = ["GrandTotal"] if "GrandTotal" in df.columns else []
                    return df[cols + inds + extras]
            
                p1 = reorder_panel(scenario1_result["pivot_panel"])
                p2 = reorder_panel(scenario2_result["pivot_panel"])
            
                sheet_p = "PanelAllocated"
                p1.to_excel(writer, sheet_name=sheet_p, startrow=0, index=False)
                start_p = p1.shape[0] + 2
                p2.to_excel(writer, sheet_name=sheet_p, startrow=start_p, index=False)
            
                # 5) Region-wise & Size-wise Totals side-by-side, in Adjusted order
                r1 = scenario1_result["region_totals"].rename(columns={
                    "PanelAllocated":"S1_Panel","FreshAllocated":"S1_Fresh","SampleTotal":"S1_Total"
                }).set_index("Region").reindex(region_order).reset_index()
            
                r2 = scenario2_result["region_totals"].rename(columns={
                    "PanelAllocated":"S2_Panel","FreshAllocated":"S2_Fresh","SampleTotal":"S2_Total"
                }).set_index("Region").reindex(region_order).reset_index()
            
                region_totals_combined = pd.merge(r1, r2, on="Region", how="outer").loc[
                    :, ["Region","S1_Panel","S2_Panel","S1_Fresh","S2_Fresh","S1_Total","S2_Total"]
                ]
            
                sz1 = scenario1_result["size_totals"].rename(columns={
                    "PanelAllocated":"S1_Panel","FreshAllocated":"S1_Fresh","SampleTotal":"S1_Total"
                }).set_index("Size").reindex(size_order).reset_index()
            
                sz2 = scenario2_result["size_totals"].rename(columns={
                    "PanelAllocated":"S2_Panel","FreshAllocated":"S2_Fresh","SampleTotal":"S2_Total"
                }).set_index("Size").reindex(size_order).reset_index()
            
                size_totals_combined = pd.merge(sz1, sz2, on="Size", how="outer").loc[
                    :, ["Size","S1_Panel","S2_Panel","S1_Fresh","S2_Fresh","S1_Total","S2_Total"]
                ]
            
                sheet_tot = "TotalsByRegionAndSize"
                region_totals_combined.to_excel(writer, sheet_name=sheet_tot, index=False, startrow=0)
                start_tot = region_totals_combined.shape[0] + 2
                size_totals_combined.to_excel(writer, sheet_name=sheet_tot, index=False, startrow=start_tot)

                  # 6) Comparison: Scenario 1 minus Scenario 2 with BaseWeight gradient
                # -------------------------------------------------------------------
                # build diff
                df1 = scenario1_result["df_combined"].reset_index()
                df2 = scenario2_result["df_combined"].reset_index()
                common = [c for c in df1.columns if c in df2.columns]
                df_diff = df1[common].copy()
                num_cols = [c for c in common if pd.api.types.is_numeric_dtype(df1[c])]
                for c in num_cols:
                    df_diff[c] = df1[c] - df2[c]
            
                # reorder diff sheet
                def reorder_diff(df):
                    df_out = df.copy()
                    cols = [c for c in ("Region","Size") if c in df_out.columns]
                    sample_cols = []
                    bw_cols     = []
                    for ind in industries_in_input:
                        s, b = f"{ind}_Sample", f"{ind}_BaseWeight"
                        if s in df_out.columns: sample_cols.append(s)
                        if b in df_out.columns: bw_cols.append(b)
                    extras = [c for c in ("GrandTotal_Sample","GrandTotal_BaseWeight") if c in df_out.columns]
                    return df_out[cols + sample_cols + bw_cols + extras]
            
                df_diff_out = reorder_diff(df_diff)
            
                sheet_cmp = "Comparison"
                df_diff_out.to_excel(writer, sheet_name=sheet_cmp, index=False)
                ws_cmp = writer.sheets[sheet_cmp]
            
                # apply gradient to the BaseWeight diff columns (exclude GrandTotal_BaseWeight)
             #   apply_color_scale(df_diff_out, row_offset=0)

            
            # rewind & download
            excel_out.seek(0)
            st.download_button(
                label="📥 Download Combined Excel",
                data=excel_out.getvalue(),
                file_name=f"{base_filename}_comparison.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("Please upload an Excel file first.")

if __name__=="__main__":
    import cvxpy as cp
    main()













