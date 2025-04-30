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

###############################################################################
# HELPER FUNCTIONS
###############################################################################
def dfs_to_html(sections, page_title="Snapshot"):
    """
    sections = [(title, dataframe_or_styler), ...]
    Returns bytes of a single HTML file, each table in a scrollable <div>.
    """
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
    parts = []
    for title, obj in sections:
        parts.append(f"<h2>{html.escape(title)}</h2>")
        parts.append('<div class="scrollbox">')
        if isinstance(obj, Styler):
            parts.append(obj.to_html())
        else:
            # Non-Styler DataFrame
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
        return int(round(n_infinity / (1 + n_infinity / N)))  # round to int here

def create_combined_table_with_totals(df_long):
    """
    Pivot table for (OptimizedSample, BaseWeight) with row/col totals.
    """
    pivot_sample = pd.pivot_table(
        df_long,
        index=["Region","Size"],
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
        index=["Region","Size"],
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

def detailed_feasibility_check(df_long,
                               total_sample,
                               min_cell_size,
                               max_cell_size,
                               max_base_weight,
                               conversion_rate,
                               dimension_mins):
    """
    Simple check if a scenario has direct conflicts, e.g. cell min > feasible max.
    """
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
        # apply min_cell_size if feasible
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
                "Population": int(pop_i),
                "LowerBound": round(lb),
                "FeasibleMax": round(fmax),
                "ShortBy": round(lb - fmax),
                "Reason":"Cell min > feasible max"
            })

        feasible_max_by_i.append(fmax)

    sum_feas= sum(feasible_max_by_i)
    if sum_feas< total_sample:
        overall_conflicts.append({
            "TotalSample": total_sample,
            "SumFeasibleMax": int(round(sum_feas)),
            "ShortBy": int(round(total_sample - sum_feas)),
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
                        "SumFeasibleMax": int(round(sum_fmax_dim)),
                        "ShortBy": int(round(req_min- sum_fmax_dim)),
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
            "SlackUsed": round(s_tot.value,0),
            "Comment":f"Sum(x) ended up {int(total_sample- s_tot.value)}, short by {round(s_tot.value,0)}"
        })
    for (dt,val), var in dim_slacks.items():
        if var.value>1e-8:
            slack_usage.append({
                "Constraint":f"DimensionMin {dt}={val}",
                "SlackUsed": round(var.value,0),
                "Comment":"We missed that dimension min by this amount"
            })
    for i,var in enumerate(cell_slacks):
        if var.value>1e-8:
            r= df_long.loc[i,"Region"]
            s= df_long.loc[i,"Size"]
            ind= df_long.loc[i,"Industry"]
            slack_usage.append({
                "Constraint":f"CellMin i={i} (R={r}, Sz={s}, Ind={ind})",
                "SlackUsed": round(var.value,0),
                "Comment":"We missed the cell-level LB by this amount"
            })

    df_slack_usage= pd.DataFrame(slack_usage).sort_values("SlackUsed", ascending=False)
    return df_slack_sol, df_slack_usage, prob.status

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

    # dimension constraints
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

def allocate_panel_fresh(df_long_sol, df_panel_wide, df_fresh_wide):
    """
    Splits final sample into PanelAllocated & FreshAllocated, no decimals.
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
# 6) MAIN APP
###############################################################################
def main():
    st.title("Survey Design")
    st.write("""
    **Features**:
    1. Two sheets: 'panel' and 'fresh'.
    2. Checkbox: if unchecked => Adjusted Universe = max(panel,fresh), if checked => sum(panel+fresh).
    3. Check single-constraint feasibility. If there's a direct conflict, show the Overall, Cell, and Dimension conflicts as tables.
    4. If still we fail at the solver => run a slack-based diagnostic to see combined conflict.
    5. Two scenarios with separate parameters, plus side-by-side comparison.
    6. Dimension minimums auto-calculated from sample-size formula (override in sidebar).
    7. HTML & Excel downloads, with no decimal places anywhere.
    """)

    # Scenario 1 params
    st.sidebar.header("Parameters for Scenario 1")
    total_sample_1= st.sidebar.number_input("Total Sample", value=1000, key="total_sample_1")
    min_cell_size_1= st.sidebar.number_input("Min Cell Size", value=4, key="min_cell_size_1")
    max_cell_size_1= st.sidebar.number_input("Max Cell Size", value=40, key="max_cell_size_1")
    max_base_weight_1= st.sidebar.number_input("Max Base Weight", value=600, key="max_base_weight_1")
    solver_choice_1= st.sidebar.selectbox("Solver (S1)", ["SCIP","ECOS_BB"], index=0, key="solver_choice_1")
    conversion_rate_1= st.sidebar.number_input("Conversion Rate (S1)", value=0.3, step=0.01, key="conv_rate_1")

    use_sum_universe = st.sidebar.checkbox("Use sum(panel,fresh) instead of max(panel,fresh)", value=False)

    st.sidebar.markdown("---")
    st.sidebar.markdown("**Sample Size Formula Inputs (Scenario 1)**")
    z_score_1= st.sidebar.number_input("Z-Score (S1)", value=1.644853627, format="%.9f", key="z_score_1")
    margin_of_error_1= st.sidebar.number_input("Margin of Error (S1)", value=0.075, format="%.3f", key="moe_1")
    p_1= st.sidebar.number_input("p (Prop) (S1)", value=0.5, format="%.2f", key="p_1")

    st.sidebar.markdown("---")

    # Scenario 2 params
    st.sidebar.header("Parameters for Scenario 2")
    total_sample_2= st.sidebar.number_input("Total Sample (S2)", value=800, key="total_sample_2")
    min_cell_size_2= st.sidebar.number_input("Min Cell Size (S2)", value=4, key="min_cell_size_2")
    max_cell_size_2= st.sidebar.number_input("Max Cell Size (S2)", value=40, key="max_cell_size_2")
    max_base_weight_2= st.sidebar.number_input("Max Base Weight (S2)", value=600, key="max_base_weight_2")
    solver_choice_2= st.sidebar.selectbox("Solver (S2)", ["SCIP","ECOS_BB"], index=1, key="solver_choice_2")
    conversion_rate_2= st.sidebar.number_input("Conversion Rate (S2)", value=0.3, step=0.01, key="conv_rate_2")

    st.sidebar.markdown("**Sample Size Formula Inputs (Scenario 2)**")
    z_score_2= st.sidebar.number_input("Z-Score (S2)", value=1.644853627, format="%.9f", key="z_score_2")
    margin_of_error_2= st.sidebar.number_input("Margin of Error (S2)", value=0.075, format="%.3f", key="moe_2")
    p_2= st.sidebar.number_input("p (Prop) (S2)", value=0.5, format="%.2f", key="p_2")

    st.sidebar.markdown("---")

    uploaded_file= st.file_uploader("Upload Excel with 'panel','fresh'", type=["xlsx"])

    if uploaded_file is not None:
        base_filename = uploaded_file.name.rsplit('.', 1)[0]

        try:
            df_panel= pd.read_excel(uploaded_file, sheet_name="panel")
            df_fresh= pd.read_excel(uploaded_file, sheet_name="fresh")
        except Exception as e:
            st.error(f"Error reading 'panel'/'fresh' => {e}")
            return

        st.subheader("Original Panel Table")
        st.data_editor(df_panel, key="orig_panel", use_container_width=True)
        st.subheader("Original Fresh Table")
        st.data_editor(df_fresh, key="orig_fresh", use_container_width=True)

        # Build Adjusted Universe
        df_adjusted = df_panel.copy()
        id_cols= ["Region","Size"]
        industry_cols= [c for c in df_panel.columns if c not in id_cols]
        for c in industry_cols:
            if c in df_fresh.columns:
                if use_sum_universe:
                    df_adjusted[c] = df_panel[c].fillna(0) + df_fresh[c].fillna(0)
                else:
                    df_adjusted[c] = np.maximum(df_panel[c].fillna(0), df_fresh[c].fillna(0))

        st.subheader("Adjusted Universe Table")
        st.data_editor(df_adjusted, key="adj_universe", use_container_width=True)

        # Gather dimension sets
        all_regions = df_adjusted["Region"].dropna().unique()
        all_sizes   = df_adjusted["Size"].dropna().unique()
        all_inds    = industry_cols

        # dimension min dicts
        dimension_mins_1= {"Region":{}, "Size":{}, "Industry":{}}
        dimension_mins_2= {"Region":{}, "Size":{}, "Industry":{}}

        # Precompute scenario 1 n_infinity
        n_infinity_1 = compute_n_infinity(z_score_1, margin_of_error_1, p_1)
        def sum_pop_in_dim_s1(dim_type, val):
            subset= df_adjusted[df_adjusted[dim_type]== val]
            return subset[industry_cols].fillna(0).sum().sum()

        # Precompute scenario 2 n_infinity
        n_infinity_2 = compute_n_infinity(z_score_2, margin_of_error_2, p_2)
        def sum_pop_in_dim_s2(dim_type, val):
            subset= df_adjusted[df_adjusted[dim_type]== val]
            return subset[industry_cols].fillna(0).sum().sum()

        # Fill dimension minimums from formula
        for r in all_regions:
            pop_1= sum_pop_in_dim_s1("Region", r)
            dimension_mins_1["Region"][r] = compute_fpc_min(pop_1, n_infinity_1)
            pop_2= sum_pop_in_dim_s2("Region", r)
            dimension_mins_2["Region"][r] = compute_fpc_min(pop_2, n_infinity_2)
        for s in all_sizes:
            pop_1= sum_pop_in_dim_s1("Size", s)
            dimension_mins_1["Size"][s] = compute_fpc_min(pop_1, n_infinity_1)
            pop_2= sum_pop_in_dim_s2("Size", s)
            dimension_mins_2["Size"][s] = compute_fpc_min(pop_2, n_infinity_2)
        for i in all_inds:
            pop_1= df_adjusted[i].fillna(0).sum()
            dimension_mins_1["Industry"][i] = compute_fpc_min(pop_1, n_infinity_1)
            pop_2= df_adjusted[i].fillna(0).sum()
            dimension_mins_2["Industry"][i] = compute_fpc_min(pop_2, n_infinity_2)

        # We'll show these in the final output, so no separate sidebar override in this example.
        # (If you want to override them, simply add number_input loops.)

        # Helper for scenario run
        def run_scenario(total_sample, min_c, max_c, max_bw, solver_choice, conv_rate, dimension_mins):
            scenario_result= {}
            try:
                df_long_final, df_cell_conf, df_dim_conf, solverinfo = run_optimization(
                    df_wide=df_adjusted,
                    total_sample=total_sample,
                    min_cell_size=min_c,
                    max_cell_size=max_c,
                    max_base_weight=max_bw,
                    solver_choice=solver_choice,
                    dimension_mins=dimension_mins,
                    conversion_rate=conv_rate
                )
                if df_long_final is None:
                    # single-constraint conflict
                    scenario_result["success"]=False
                    scenario_result["cell_conflicts"]= df_cell_conf
                    scenario_result["dim_conflicts"] = df_dim_conf
                    scenario_result["overall_conflicts"]= solverinfo
                else:
                    # success
                    scenario_result["success"]=True
                    scenario_result["solverinfo"]= solverinfo
                    # allocate
                    df_alloc= allocate_panel_fresh(df_long_final, df_panel, df_fresh)
                    scenario_result["df_alloc"] = df_alloc
                    # Build pivot etc.
                    region_totals = (
                        df_alloc
                        .groupby("Region")[["PanelAllocated","FreshAllocated"]]
                        .sum()
                        .assign(SampleTotal=lambda d: d["PanelAllocated"]+ d["FreshAllocated"])
                        .reset_index()
                    )

                    size_totals = (
                        df_alloc
                        .groupby("Size")[["PanelAllocated","FreshAllocated"]]
                        .sum()
                        .assign(SampleTotal=lambda d: d["PanelAllocated"]+ d["FreshAllocated"])
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
                        margins_name="GrandTotal"
                    ).reset_index()

                    pivot_fresh = pd.pivot_table(
                        df_alloc,
                        index=["Region","Size"],
                        columns="Industry",
                        values="FreshAllocated",
                        aggfunc='sum',
                        fill_value=0,
                        margins=True,
                        margins_name="GrandTotal"
                    ).reset_index()

                    df_combined = create_combined_table_with_totals(df_alloc)
                    # reorder
                    id_cols= [c for c in ["Region","Size"] if c in df_combined.columns]
                    gt_samp= [c for c in df_combined.columns if c=="GrandTotal_Sample"]
                    gt_bw  = [c for c in df_combined.columns if c=="GrandTotal_BaseWeight"]
                    s_cols = [c for c in df_combined.columns if c.endswith("_Sample") and c not in gt_samp]
                    bw_cols= [c for c in df_combined.columns if c.endswith("_BaseWeight") and c not in gt_bw]
                    df_combined= df_combined[id_cols + gt_samp + s_cols + bw_cols + gt_bw]

                    # Round everything to 0 decimals, cast to int, even base weights
                    numeric_only= df_combined.select_dtypes(include=[np.number]).columns
                    df_combined[numeric_only]= df_combined[numeric_only].round(0).astype(int, errors="ignore")

                    # color scale
                    norm_bw_cols= [c for c in bw_cols if c!="GrandTotal_BaseWeight"]
                    if len(df_combined)>1 and norm_bw_cols:
                        norm_df = df_combined.iloc[:-1]  # exclude total row
                        global_min= norm_df[norm_bw_cols].min().min()
                        global_max= norm_df[norm_bw_cols].max().max()
                        global_mid= np.percentile(norm_df[norm_bw_cols].stack(), 50)
                    else:
                        global_min= 0
                        global_mid= 0
                        global_max= 0
                    custom_cmap= LinearSegmentedColormap.from_list("custom", ["#00FF00","#FFFF00","#FF0000"])
                    norm_obj= TwoSlopeNorm(vmin=global_min, vcenter=global_mid, vmax=global_max)
                    def baseweight_color(v):
                        val= norm_obj(v)
                        return f"background-color: {mcolors.to_hex(custom_cmap(val))}"
                    def style_bwcol(series):
                        n= len(series)
                        return [baseweight_color(x) if i<n-1 else "" for i,x in enumerate(series)]
                    stcol= df_combined.style.apply(style_bwcol, subset=norm_bw_cols)

                    # population pivot
                    pivot_pop= pd.pivot_table(
                        df_alloc,
                        index=["Region","Size"],
                        columns="Industry",
                        values="Population",
                        aggfunc='sum',
                        fill_value=0,
                        margins=True,
                        margins_name="GrandTotal"
                    ).reset_index()

                    # prop pivot
                    pivot_propsample= pd.pivot_table(
                        df_alloc,
                        index=["Region","Size"],
                        columns="Industry",
                        values="PropSample",
                        aggfunc='mean',
                        fill_value=0
                    ).reset_index()
                    numeric_cols= pivot_propsample.select_dtypes(include="number").columns
                    pivot_propsample[numeric_cols]= pivot_propsample[numeric_cols].round(0).astype(int)

                    # store
                    scenario_result["region_totals"]= region_totals
                    scenario_result["size_totals"]= size_totals
                    scenario_result["pivot_panel"]= pivot_panel.round(0).astype(int, errors="ignore")
                    scenario_result["pivot_fresh"]= pivot_fresh.round(0).astype(int, errors="ignore")
                    scenario_result["df_combined"]= df_combined
                    scenario_result["df_combined_style"]= stcol
                    scenario_result["pivot_pop"]= pivot_pop
                    scenario_result["pivot_propsample"]= pivot_propsample.round(0).astype(int, errors="ignore")

            except ValueError as e:
                if "No solver found a feasible solution" in str(e):
                    # slack diag
                    df_longX= df_adjusted.melt(
                        id_vars=["Region","Size"],
                        var_name="Industry",
                        value_name="Population"
                    ).fillna(0)
                    diag_sol, diag_usage, diag_status= diagnose_infeasibility_slacks(
                        df_longX,
                        total_sample,
                        min_c,
                        max_c,
                        max_bw,
                        dimension_mins,
                        conv_rate
                    )
                    scenario_result["success"]=False
                    scenario_result["error_msg"]= str(e)
                    scenario_result["diag_sol"]= diag_sol
                    scenario_result["diag_usage"]= diag_usage
                    scenario_result["diag_status"]= diag_status
                else:
                    scenario_result["success"]=False
                    scenario_result["error_msg"]= str(e)
            except Exception as e2:
                scenario_result["success"]=False
                scenario_result["error_msg"]= f"Solver Error: {e2}"

            return scenario_result

        if st.button("Run Optimization for Both Scenarios"):
            # Create param data frames
            params_1 = pd.DataFrame([
                ("Total Sample", total_sample_1),
                ("Min Cell Size", min_cell_size_1),
                ("Max Cell Size", max_cell_size_1),
                ("Max Base Weight", max_base_weight_1),
                ("Solver", solver_choice_1),
                ("Conversion Rate", conversion_rate_1),
                ("Z-Score", z_score_1),
                ("Margin of Error", margin_of_error_1),
                ("p (Proportion)", p_1),
                ("Use sum(panel,fresh)?", use_sum_universe)
            ], columns=["Parameter","Value"])

            params_2 = pd.DataFrame([
                ("Total Sample", total_sample_2),
                ("Min Cell Size", min_cell_size_2),
                ("Max Cell Size", max_cell_size_2),
                ("Max Base Weight", max_base_weight_2),
                ("Solver", solver_choice_2),
                ("Conversion Rate", conversion_rate_2),
                ("Z-Score", z_score_2),
                ("Margin of Error", margin_of_error_2),
                ("p (Proportion)", p_2),
                ("Use sum(panel,fresh)?", use_sum_universe)
            ], columns=["Parameter","Value"])

            # Convert dimension mins to DataFrame
            def dims_to_df(dims_dict, label):
                rows= []
                for dt, val_dict in dims_dict.items():
                    for v, m in val_dict.items():
                        rows.append((dt, v, m))
                return pd.DataFrame(rows, columns=["DimensionType","DimensionValue","MinSample"]).assign(Scenario=label)
            dimdf_1= dims_to_df(dimension_mins_1, label="Scenario1")
            dimdf_2= dims_to_df(dimension_mins_2, label="Scenario2")

            scenario1_result= run_scenario(
                total_sample_1, min_cell_size_1, max_cell_size_1, max_base_weight_1,
                solver_choice_1, conversion_rate_1, dimension_mins_1
            )
            scenario2_result= run_scenario(
                total_sample_2, min_cell_size_2, max_cell_size_2, max_base_weight_2,
                solver_choice_2, conversion_rate_2, dimension_mins_2
            )

            st.header("Scenario 1 Results")
            if not scenario1_result["success"]:
                # conflicts or error
                if "cell_conflicts" in scenario1_result:
                    st.error("Single-constraint conflict(s) found (Scenario 1)")
                    oc= scenario1_result["overall_conflicts"]
                    if oc is not None and not oc.empty:
                        st.subheader("Overall Conflicts (S1)")
                        st.data_editor(oc, key="s1_overall_conf", use_container_width=True)
                    cc= scenario1_result["cell_conflicts"]
                    if cc is not None and not cc.empty:
                        st.subheader("Cell Conflicts (S1)")
                        st.data_editor(cc, key="s1_cell_conf", use_container_width=True)
                    dc= scenario1_result["dim_conflicts"]
                    if dc is not None and not dc.empty:
                        st.subheader("Dimension Conflicts (S1)")
                        st.data_editor(dc, key="s1_dim_conf", use_container_width=True)
                elif "error_msg" in scenario1_result:
                    st.error(scenario1_result["error_msg"])
                    if "diag_sol" in scenario1_result:
                        st.warning("Slack-based diagnostic (S1)")
                        st.dataframe(scenario1_result["diag_sol"], key="s1_diag_sol")
                        st.dataframe(scenario1_result["diag_usage"], key="s1_diag_usage")
                        st.write(f"Status: {scenario1_result['diag_status']}")
            else:
                # success
                st.success(f"Solved with solver: {scenario1_result['solverinfo']}")
                st.subheader("Panel Allocation (Scenario 1)")
                st.data_editor(scenario1_result["pivot_panel"], key="s1_panel", use_container_width=True)
                st.subheader("Fresh Allocation (Scenario 1)")
                st.data_editor(scenario1_result["pivot_fresh"], key="s1_fresh", use_container_width=True)
                st.subheader("Allocated Sample & Base Weights (Scenario 1)")
                st.dataframe(scenario1_result["df_combined_style"], key="s1_combined")
                st.subheader("Region Totals (Scenario 1)")
                st.data_editor(scenario1_result["region_totals"], key="s1_region_totals", use_container_width=True)
                st.subheader("Size Totals (Scenario 1)")
                st.data_editor(scenario1_result["size_totals"], key="s1_size_totals", use_container_width=True)

            st.header("Scenario 2 Results")
            if not scenario2_result["success"]:
                # conflict
                if "cell_conflicts" in scenario2_result:
                    st.error("Single-constraint conflict(s) found (Scenario 2)")
                    oc= scenario2_result["overall_conflicts"]
                    if oc is not None and not oc.empty:
                        st.subheader("Overall Conflicts (S2)")
                        st.data_editor(oc, key="s2_overall_conf", use_container_width=True)
                    cc= scenario2_result["cell_conflicts"]
                    if cc is not None and not cc.empty:
                        st.subheader("Cell Conflicts (S2)")
                        st.data_editor(cc, key="s2_cell_conf", use_container_width=True)
                    dc= scenario2_result["dim_conflicts"]
                    if dc is not None and not dc.empty:
                        st.subheader("Dimension Conflicts (S2)")
                        st.data_editor(dc, key="s2_dim_conf", use_container_width=True)
                elif "error_msg" in scenario2_result:
                    st.error(scenario2_result["error_msg"])
                    if "diag_sol" in scenario2_result:
                        st.warning("Slack-based diagnostic (S2)")
                        st.dataframe(scenario2_result["diag_sol"], key="s2_diag_sol")
                        st.dataframe(scenario2_result["diag_usage"], key="s2_diag_usage")
                        st.write(f"Status: {scenario2_result['diag_status']}")
            else:
                # success
                st.success(f"Solved with solver: {scenario2_result['solverinfo']}")
                st.subheader("Panel Allocation (Scenario 2)")
                st.data_editor(scenario2_result["pivot_panel"], key="s2_panel", use_container_width=True)
                st.subheader("Fresh Allocation (Scenario 2)")
                st.data_editor(scenario2_result["pivot_fresh"], key="s2_fresh", use_container_width=True)
                st.subheader("Allocated Sample & Base Weights (Scenario 2)")
                st.dataframe(scenario2_result["df_combined_style"], key="s2_combined")
                st.subheader("Region Totals (Scenario 2)")
                st.data_editor(scenario2_result["region_totals"], key="s2_region_totals", use_container_width=True)
                st.subheader("Size Totals (Scenario 2)")
                st.data_editor(scenario2_result["size_totals"], key="s2_size_totals", use_container_width=True)

            # ---------- Side-by-side table -------------
            st.header("Side-by-Side Comparison Table (Samples & Base Weights)")
            # only if both scenarios succeeded
            if scenario1_result.get("success") and scenario2_result.get("success"):
                df_s1 = scenario1_result["df_alloc"].copy()
                df_s2 = scenario2_result["df_alloc"].copy()
                # keep Region/Size/Industry + sample/baseweight
                s1_cols = ["Region","Size","Industry","OptimizedSample","BaseWeight"]
                s2_cols = ["Region","Size","Industry","OptimizedSample","BaseWeight"]
                df_s1 = df_s1[s1_cols]
                df_s2 = df_s2[s2_cols]
                merged = pd.merge(
                    df_s1, df_s2,
                    on=["Region","Size","Industry"],
                    how="outer",
                    suffixes=("_S1","_S2")
                )
                # difference
                merged["SampleDiff"]  = merged["OptimizedSample_S1"] - merged["OptimizedSample_S2"]
                merged["BaseWDiff"]   = merged["BaseWeight_S1"] - merged["BaseWeight_S2"]
                # round (which is typically integer anyway)
                numeric_cols = merged.select_dtypes(include="number").columns
                merged[numeric_cols] = merged[numeric_cols].round(0).astype(int, errors="ignore")

                st.data_editor(merged, key="side_by_side_alloc", use_container_width=True)

            # ---------- Download (HTML & Excel) ----------
            scenario_sections = []
            scenario_sections.append(("Parameters (Scenario 1)", params_1))
            scenario_sections.append(("Parameters (Scenario 2)", params_2))
            scenario_sections.append(("Dimension Minimums (Scenario 1)", dimdf_1))
            scenario_sections.append(("Dimension Minimums (Scenario 2)", dimdf_2))

            # If scenario1 succeeded
            if scenario1_result.get("success"):
                scenario_sections.append(("Scenario 1: Allocated Sample & Base Weights", scenario1_result["df_combined_style"]))

            # If scenario2 succeeded
            if scenario2_result.get("success"):
                scenario_sections.append(("Scenario 2: Allocated Sample & Base Weights", scenario2_result["df_combined_style"]))

            # Side-by-side table
            if scenario1_result.get("success") and scenario2_result.get("success"):
                scenario_sections.append(("Side-by-Side Comparison Table", merged))

            html_bytes = dfs_to_html(scenario_sections, page_title=base_filename)
            st.download_button(
                label="Download HTML",
                data=html_bytes,
                file_name=f"{base_filename}_comparison.html",
                mime="text/html"
            )

            # Excel
            excel_out= io.BytesIO()
            with pd.ExcelWriter(excel_out, engine="openpyxl") as writer:
                # 1) Scenario 1 parameters
                params_1.to_excel(writer, sheet_name="S1_Parameters", index=False)
                # 2) Scenario 2 parameters
                params_2.to_excel(writer, sheet_name="S2_Parameters", index=False)
                # 3) Dim mins
                dimdf_1.to_excel(writer, sheet_name="S1_DimMins", index=False)
                dimdf_2.to_excel(writer, sheet_name="S2_DimMins", index=False)

                def write_scenario_tables(scenario_result, label):
                    # region panel/fresh, combined, population, propsample
                    # We'll do 1 sheet each for: panel, fresh, combined, pop, propsample
                    df_panel = scenario_result["pivot_panel"]
                    df_fresh = scenario_result["pivot_fresh"]
                    df_comb  = scenario_result["df_combined"]
                    pop      = scenario_result["pivot_pop"]
                    propsamp = scenario_result["pivot_propsample"]

                    df_panel.to_excel(writer, sheet_name=f"{label}_Panel", index=False)
                    df_fresh.to_excel(writer, sheet_name=f"{label}_Fresh", index=False)
                    df_comb.to_excel(writer, sheet_name=f"{label}_SampleBase", index=False)
                    pop.to_excel(writer, sheet_name=f"{label}_Population", index=False)
                    propsamp.to_excel(writer, sheet_name=f"{label}_PropSample", index=False)

                # scenario 1
                if scenario1_result.get("success"):
                    write_scenario_tables(scenario1_result, "S1")

                # scenario 2
                if scenario2_result.get("success"):
                    write_scenario_tables(scenario2_result, "S2")

                # side by side
                if scenario1_result.get("success") and scenario2_result.get("success"):
                    merged.to_excel(writer, sheet_name="S1vsS2_SideBySide", index=False)

            excel_out.seek(0)
            st.download_button(
                label="Download Excel",
                data=excel_out.getvalue(),
                file_name=f"{base_filename}_comparison.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    else:
        st.warning("Please upload an Excel file first.")

if __name__=="__main__":
    main()
