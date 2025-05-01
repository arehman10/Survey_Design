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
    feasible_max_by_i = []
    cell_conflicts = []
    dimension_conflicts = []
    overall_conflicts = []

    for i in range(n_cells):
        pop_i = df_long.loc[i,"Population"]
        if pop_i<=0:
            fmax=0
        else:
            conv_ub= math.ceil(pop_i*conversion_rate)
            fmax= min(pop_i, max_cell_size, conv_ub)

        lb_bw= pop_i/max_base_weight
        if fmax>= min_cell_size and pop_i>0:
            lb= max(lb_bw, min_cell_size, 0)
        else:
            lb= max(lb_bw, 0)

        if lb> fmax:
            cell_conflicts.append({
                "CellIndex": i,
                "Region": df_long.loc[i,"Region"],
                "Size": df_long.loc[i,"Size"],
                "Industry": df_long.loc[i,"Industry"],
                "Population": pop_i,
                "LowerBound": lb,
                "FeasibleMax": fmax,
                "ShortBy": lb - fmax,
                "Reason":"Cell min > feasible max"
            })

        feasible_max_by_i.append(fmax)

    sum_feas= sum(feasible_max_by_i)
    if sum_feas< total_sample:
        overall_conflicts.append({
            "TotalSample": total_sample,
            "SumFeasibleMax": sum_feas,
            "ShortBy": total_sample - sum_feas,
            "Reason":"Overall sum(feasible max) < total_sample"
        })

    # dimension checks
    dim_idx= {"Region":{}, "Size":{}, "Industry":{}}
    for dt in dim_idx:
        for val in df_long[dt].unique():
            idx_list= df_long.index[df_long[dt]== val].tolist()
            dim_idx[dt][val]= idx_list

    for dt, val_dict in dimension_mins.items():
        for val, req_min in val_dict.items():
            if req_min>0:
                idx_list= dim_idx[dt].get(val, [])
                sum_fmax_dim= sum(feasible_max_by_i[i] for i in idx_list)
                if sum_fmax_dim< req_min:
                    dimension_conflicts.append({
                        "DimType": dt,
                        "DimName": val,
                        "RequiredMin": req_min,
                        "SumFeasibleMax": sum_fmax_dim,
                        "ShortBy": req_min- sum_fmax_dim,
                        "Reason":"Dimension min> sum feasible"
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
    If run_optimization fails, we do a slack-based approach 
    to see how far we are from feasibility.
    """
    import cvxpy as cp
    x= cp.Variable(len(df_long), nonneg=True)
    s_tot= cp.Variable(nonneg=True)
    constraints= [cp.sum(x)+ s_tot== total_sample]

    dim_slacks= {}
    dim_idx= {"Region":{}, "Size":{}, "Industry":{}}
    for dt in dim_idx:
        for val in df_long[dt].unique():
            idx_list= df_long.index[df_long[dt]== val].tolist()
            dim_idx[dt][val]= idx_list

    for dt, val_dict in dimension_mins.items():
        for val, req_min in val_dict.items():
            if req_min>0:
                sd= cp.Variable(nonneg=True)
                dim_slacks[(dt,val)] = sd
                idx_list= dim_idx[dt][val]
                constraints.append( cp.sum(x[idx_list]) + sd>= req_min)

    cell_slacks= []
    for i in range(len(df_long)):
        pop_i= df_long.loc[i,"Population"]
        if pop_i<=0:
            fmax=0
        else:
            conv_ub= math.ceil(pop_i*conversion_rate)
            fmax= min(pop_i, max_cell_size, conv_ub)
        lb_bw= pop_i/ max_base_weight
        if fmax>= min_cell_size and pop_i>0:
            lb= max(lb_bw, min_cell_size, 0)
        else:
            lb= max(lb_bw,0)
        s_cell= cp.Variable(nonneg=True)
        cell_slacks.append(s_cell)
        constraints.append( x[i]+ s_cell>= lb )

    slack_vars= [s_tot]+ list(dim_slacks.values())+ cell_slacks
    obj= cp.Minimize( cp.sum(slack_vars) )
    prob= cp.Problem(obj, constraints)
    try:
        prob.solve(solver="ECOS", verbose=False)
    except Exception as e:
        return None, None, f"Diagnostic solver error: {e}"

    if prob.status in ["infeasible","unbounded"]:
        return None, None, f"Diagnostic problem is {prob.status}."

    x_sol= x.value
    df_slack_sol= df_long.copy()
    df_slack_sol["SlackSolution_x"]= x_sol

    slack_usage=[]
    if s_tot.value>1e-8:
        slack_usage.append({
            "Constraint":"TotalSample",
            "SlackUsed": s_tot.value,
            "Comment":f"Sum(x) ended up {total_sample- s_tot.value}, short by {s_tot.value}"
        })
    for (dt,val), var in dim_slacks.items():
        if var.value>1e-8:
            slack_usage.append({
                "Constraint":f"DimensionMin {dt}={val}",
                "SlackUsed": var.value,
                "Comment":"We missed that dimension min by this amount"
            })
    for i,var in enumerate(cell_slacks):
        if var.value>1e-8:
            r= df_long.loc[i,"Region"]
            s= df_long.loc[i,"Size"]
            ind= df_long.loc[i,"Industry"]
            slack_usage.append({
                "Constraint":f"CellMin i={i} (R={r}, Sz={s}, Ind={ind})",
                "SlackUsed": var.value,
                "Comment":"We missed the cell-level LB by this amount"
            })

    df_slack_usage= pd.DataFrame(slack_usage).sort_values("SlackUsed", ascending=False)
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
                     conversion_rate):
    """
    Convert wide->long, do feasibility check, solve MIP, return solution or None if direct conflict.
    """
    import cvxpy as cp

    id_cols= ["Region","Size"]
    data_cols= [c for c in df_wide.columns if c not in id_cols]
    df_long= df_wide.melt(
        id_vars= id_cols,
        value_vars=data_cols,
        var_name="Industry",
        value_name="Population"
    ).reset_index(drop=True)
    df_long["Population"]= df_long["Population"].fillna(0)

    total_pop= df_long["Population"].sum()
    if total_pop>0:
        df_long["PropSample"]= df_long["Population"]*(total_sample/ total_pop)
    else:
        df_long["PropSample"]= 0

    # Quick single-constraint check
    df_overall, df_cells, df_dims = detailed_feasibility_check(
        df_long, total_sample, min_cell_size, max_cell_size, max_base_weight, 
        conversion_rate, dimension_mins
    )
    if (not df_overall.empty) or (not df_cells.empty) or (not df_dims.empty):
        # Return None to show conflict
        return None, df_cells, df_dims, df_overall

    n_cells= len(df_long)
    x= cp.Variable(n_cells, integer=True)
    obj_dev= cp.sum_squares(x - df_long["PropSample"])
    objective= cp.Minimize(obj_dev)

    constraints= [cp.sum(x)== total_sample]
    for i in range(n_cells):
        pop_i= df_long.loc[i,"Population"]
        if pop_i<=0:
            constraints.append(x[i]==0)
            continue
        conv_ub= math.ceil(pop_i* conversion_rate)
        fmax= min(pop_i, max_cell_size, conv_ub)
        lb_bw= pop_i/ max_base_weight
        if fmax>= min_cell_size and pop_i>0:
            lb= max(lb_bw, min_cell_size, 0)
        else:
            lb= max(lb_bw,0)
        constraints.append(x[i]>= lb)
        constraints.append(x[i]<= fmax)
        constraints.append(x[i]<= pop_i)

    dim_idx= {"Region":{}, "Size":{}, "Industry":{}}
    for dt in dim_idx:
        for val in df_long[dt].unique():
            idx_list= df_long.index[df_long[dt]== val].tolist()
            dim_idx[dt][val]= idx_list

    for dt, val_dict in dimension_mins.items():
        for val, req_min in val_dict.items():
            if req_min>0 and val in dim_idx[dt]:
                idx_list= dim_idx[dt][val]
                constraints.append(cp.sum(x[idx_list])>= req_min)

    problem= cp.Problem(objective, constraints)
    candidate_solvers= ["SCIP","ECOS_BB"]
    if solver_choice in candidate_solvers:
        chosen_solvers= [solver_choice]+ [s for s in candidate_solvers if s!= solver_choice]
    else:
        chosen_solvers= candidate_solvers

    solution_found= False
    solver_that_succeeded= None
    last_error= None
    for s in chosen_solvers:
        try:
            res= problem.solve(solver=s)
            if problem.status not in ["infeasible","unbounded"] and x.value is not None:
                solution_found= True
                solver_that_succeeded= s
                break
        except Exception as e:
            last_error= e

    if not solution_found:
        raise ValueError(f"No solver found a feasible solution among {chosen_solvers}. Last error was: {last_error}")

    x_sol= np.round(x.value).astype(int)
    df_long["OptimizedSample"]= x_sol

    def basew(row):
        if row["OptimizedSample"]>0:
            return row["Population"] / row["OptimizedSample"]
        else:
            return np.nan
    df_long["BaseWeight"]= df_long.apply(basew, axis=1)

    return df_long, None, None, solver_that_succeeded

###############################################################################
# 4) ALLOCATE BETWEEN PANEL & FRESH
###############################################################################
def allocate_panel_fresh(df_long_sol, df_panel_wide, df_fresh_wide):
    """
    Splits final sample into PanelAllocated & FreshAllocated.
    """
    identifier_cols= ["Region","Size"]
    panel_inds= [c for c in df_panel_wide.columns if c not in identifier_cols]

    df_panel_long= df_panel_wide.melt(
        id_vars=identifier_cols,
        value_vars= panel_inds,
        var_name="Industry",
        value_name="PanelPop"
    ).reset_index(drop=True)

    panel_dict= {}
    for rowi in df_panel_long.itertuples(index=False):
        key= (rowi.Region, rowi.Size, rowi.Industry)
        val= rowi.PanelPop if (rowi.PanelPop is not None and not pd.isna(rowi.PanelPop)) else 0
        panel_dict[key]= val

    df_long_sol["PanelAllocated"]= 0
    df_long_sol["FreshAllocated"]= 0

    for i in range(len(df_long_sol)):
        reg= df_long_sol.loc[i,"Region"]
        sz= df_long_sol.loc[i,"Size"]
        ind= df_long_sol.loc[i,"Industry"]
        x_i= df_long_sol.loc[i,"OptimizedSample"]
        if x_i<=0:
            continue
        panelPop= panel_dict.get((reg,sz,ind), 0)
        half= x_i//2
        leftover= x_i % 2
        panel_alloc= half
        if leftover==1:
            if panel_alloc+1<= panelPop:
                panel_alloc= panel_alloc+1
            else:
                if panel_alloc> panelPop:
                    panel_alloc= panelPop
        if panel_alloc> panelPop:
            panel_alloc= panelPop

        fresh_alloc= x_i - panel_alloc
        df_long_sol.at[i,"PanelAllocated"]= panel_alloc
        df_long_sol.at[i,"FreshAllocated"]= fresh_alloc

    return df_long_sol

###############################################################################
# 5) MAIN APP (SAME STRUCTURE, ADDED SAVE/LOAD SESSION)
###############################################################################
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

        with st.expander("Original Panel Table"):
            st.data_editor(df_panel_wide, use_container_width=True, key="wide_panel_scenario")
        with st.expander("Original Fresh Table"):
            st.data_editor(df_fresh_wide, use_container_width=True, key="wide_fresh_scenario")

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

        with st.sidebar.expander("Dimension Minimum Overrides (Auto-Calculated, Then Overridable)", expanded=True):
            st.write("### Scenario 1")
            for r in all_regions:
                pop_ = sum_pop_in_dim_1(df_adjusted,"Region", r)
                defMin= compute_fpc_min(pop_, n_infinity_1)
                user_val= st.number_input(f"Min sample for Region={r} (S1)", 
                                          min_value=0, value=int(round(defMin)), step=1, key=f"S1_reg_{r}")
                dimension_mins_1["Region"][r]= user_val
            for sz in all_sizes:
                pop_ = sum_pop_in_dim_1(df_adjusted,"Size", sz)
                defMin= compute_fpc_min(pop_, n_infinity_1)
                user_val= st.number_input(f"Min sample for Size={sz} (S1)",
                                          min_value=0, value=int(round(defMin)), step=1, key=f"S1_size_{sz}")
                dimension_mins_1["Size"][sz]= user_val
            for ind_ in all_inds:
                pop_= df_adjusted[ind_].fillna(0).sum()
                defMin= compute_fpc_min(pop_, n_infinity_1)
                user_val= st.number_input(f"Min sample for Industry={ind_} (S1)",
                                          min_value=0, value=int(round(defMin)), step=1, key=f"S1_ind_{ind_}")
                dimension_mins_1["Industry"][ind_]= user_val

            st.write("### Scenario 2")
            for r in all_regions:
                pop_ = sum_pop_in_dim_2(df_adjusted,"Region", r)
                defMin= compute_fpc_min(pop_, n_infinity_2)
                user_val= st.number_input(f"Min sample for Region={r} (S2)", 
                                          min_value=0, value=int(round(defMin)), step=1, key=f"S2_reg_{r}")
                dimension_mins_2["Region"][r]= user_val
            for sz in all_sizes:
                pop_ = sum_pop_in_dim_2(df_adjusted,"Size", sz)
                defMin= compute_fpc_min(pop_, n_infinity_2)
                user_val= st.number_input(f"Min sample for Size={sz} (S2)",
                                          min_value=0, value=int(round(defMin)), step=1, key=f"S2_size_{sz}")
                dimension_mins_2["Size"][sz]= user_val
            for ind_ in all_inds:
                pop_= df_adjusted[ind_].fillna(0).sum()
                defMin= compute_fpc_min(pop_, n_infinity_2)
                user_val= st.number_input(f"Min sample for Industry={ind_} (S2)",
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
                        conversion_rate=conv_rate
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
                        df_alloc = allocate_panel_fresh(df_long_final, df_panel_wide, df_fresh_wide)
                        scenario_result["df_alloc"] = df_alloc

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

                st.subheader("Fresh (Scenario 1)")
                st.data_editor(scenario1_result["pivot_fresh"], use_container_width=True, key="s1_fresh")
                st.subheader("Panel (Scenario 1)")
                st.data_editor(scenario1_result["pivot_panel"], use_container_width=True, key="s1_panel")
                
                st.subheader("Allocated Sample & Base Weights (Scenario 1)")
                st.dataframe(scenario1_result["df_combined_style"], key="s1_combined")

                st.subheader("Region-wise Sample Totals (Scenario 1)")
                st.data_editor(scenario1_result["region_totals"], use_container_width=True, key="s1_region_totals")

                st.subheader("Size-wise Sample Totals (Scenario 1)")
                st.data_editor(scenario1_result["size_totals"], use_container_width=True, key="s1_size_totals")


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
               
                st.subheader("Fresh (Scenario 2)")
                st.data_editor(scenario2_result["pivot_fresh"], use_container_width=True, key="s2_fresh")
                st.subheader("Panel (Scenario 2)")
                st.data_editor(scenario2_result["pivot_panel"], use_container_width=True, key="s2_panel")

                
                st.subheader("Allocated Sample & Base Weights (Scenario 2)")
                st.dataframe(scenario2_result["df_combined_style"], key="s2_combined")

                st.subheader("Region-wise Sample Totals (Scenario 2)")
                st.data_editor(scenario2_result["region_totals"], use_container_width=True, key="s2_region_totals")

                st.subheader("Size-wise Sample Totals (Scenario 2)")
                st.data_editor(scenario2_result["size_totals"], use_container_width=True, key="s2_size_totals")

            

            # -------------- COMPARISON --------------
            st.header("Comparison: Scenario 1 minus Scenario 2")
            if scenario1_result.get("success") and scenario2_result.get("success"):
                dfc1 = scenario1_result["df_combined"].copy().reset_index(drop=True)
                dfc2 = scenario2_result["df_combined"].copy().reset_index(drop=True)
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

            
            # For Excel, we combine scenario1 & scenario2 sheets in *one* workbook
            excel_out = io.BytesIO()
            with pd.ExcelWriter(excel_out, engine="openpyxl") as writer:
                 # 1) Single "Adjusted_Universe" sheet
                df_adjusted.to_excel(writer, sheet_name="Adjusted_Universe", index=False)
            
                # 2) Scenario 1 => param/min sheet + allocated baseweight
                if scenario1_result.get("success"):
                    # a) param/min
                    s1_params_sheet = "S1_ParametersAndMins"
                    params_df.to_excel(writer, sheet_name=s1_params_sheet, index=False)
                    row_offset = params_df.shape[0] + 2
            
                    region_min_df.to_excel(writer, sheet_name=s1_params_sheet,
                                           index=False, startrow=row_offset)
                    row_offset += region_min_df.shape[0] + 2
            
                    size_min_df.to_excel(writer, sheet_name=s1_params_sheet,
                                         index=False, startrow=row_offset)
                    row_offset += size_min_df.shape[0] + 2
            
                    industry_min_df.to_excel(writer, sheet_name=s1_params_sheet,
                                             index=False, startrow=row_offset)
            
                    # b) allocated sample & baseweight
                    reorder_and_color_scenario(scenario1_result["df_combined"], "S1_Sample_with_baseweight", writer)
            
                # 3) Scenario 2 => param/min sheet + allocated baseweight
                if scenario2_result.get("success"):
                    s2_params_sheet = "S2_ParametersAndMins"
                    params_df2.to_excel(writer, sheet_name=s2_params_sheet, index=False)
                    ro2 = params_df2.shape[0] + 2
            
                    region_min_df2.to_excel(writer, sheet_name=s2_params_sheet,
                                            index=False, startrow=ro2)
                    ro2 += region_min_df2.shape[0] + 2
            
                    size_min_df2.to_excel(writer, sheet_name=s2_params_sheet,
                                          index=False, startrow=ro2)
                    ro2 += size_min_df2.shape[0] + 2
            
                    industry_min_df2.to_excel(writer, sheet_name=s2_params_sheet,
                                              index=False, startrow=ro2)
            
                    reorder_and_color_scenario(scenario2_result["df_combined"], "S2_Sample_with_baseweight", writer)
            
                # 4) If both => difference sheet, includes region & size
                if scenario1_result.get("success") and scenario2_result.get("success"):
                    diff_sheet = writer.book.create_sheet("ScenarioDiff")
                    # reorder to ensure region/size first, then columns in input order
                    df_diff_out = reorder_diff_sheet(df_diff)
                    df_diff_out = df_diff_out.reset_index(drop=True)
                    col_headers = list(df_diff_out.columns)
                    diff_sheet.append(col_headers)
                    for rowvals in df_diff_out.values:
                        diff_sheet.append(list(rowvals))
            
            excel_out.seek(0)
            st.download_button(
                label="Download Combined Excel (Both Scenarios)",
                data=excel_out.getvalue(),
                file_name=f"{base_filename}_comparison.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("Please upload an Excel file first.")

if __name__=="__main__":
    import cvxpy as cp
    main()
