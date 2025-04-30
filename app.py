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
import streamlit.components.v1 as components
from datetime import datetime
import textwrap, html
import base64, tempfile
from pandas.io.formats.style import Styler

# If using GLPK:
import swiglpk as glpk

###############################################################################
# 1) HELPER FUNCTIONS
###############################################################################
def add_grand_total_row(df, key_col_name):
    """Return df with an extra row that holds the column totals."""
    grand = df[["PanelSample", "FreshSample", "SampleTotal"]].sum().to_frame().T
    grand[key_col_name] = "Grand Total"
    # keep the same column order
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
    if N <= 0:
        return 0
    else:
        return n_infinity / (1 + n_infinity / N)

def create_combined_table_with_totals(df_long):
    """
    Pivot table for (OptimizedSample, BaseWeight) with row/col totals (margins=True).
    Honors the categorical ordering in df_long["Region","Size","Industry"] if set.
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
        sort=False  # preserve categorical order
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
        bw_col = f"{ind}_BaseWeight"
        if ind in pivot_sample.columns:
            combined[sample_col] = pivot_sample[ind]
        else:
            combined[sample_col] = np.nan

        if ind in pivot_bw.columns:
            combined[bw_col] = pivot_bw[ind]
        else:
            combined[bw_col] = np.nan

    # reorder so that "GrandTotal_Sample" & "GrandTotal_BaseWeight" appear first if they exist
    col_list = list(combined.columns)
    if "GrandTotal_Sample" in col_list:
        col_list.remove("GrandTotal_Sample")
        col_list.insert(0, "GrandTotal_Sample")
    if "GrandTotal_BaseWeight" in col_list:
        col_list.remove("GrandTotal_BaseWeight")
        col_list.insert(0, "GrandTotal_BaseWeight")

    combined = combined[col_list]
    return combined

def write_excel_combined_table(df_combined, pivot_population, pivot_propsample,
                               scenario_label=""):
    """
    Writes:
      - 'Allocated Sample' (just the sample part)
      - 'Proportional Sample'
      - 'Sample_with_baseweight' (the combined table with color-scale)
      - 'Population'
    Each sheet is prefixed with scenario_label if provided.
    """
    df_out = df_combined.reset_index()          # Region, Size become columns
    n_rows = df_out.shape[0]                    # incl. grand-total row

    # Reorder columns: region/size, then sample columns, then base-weight columns
    sample_cols = [c for c in df_out.columns if c.endswith("_Sample")]
    bw_cols     = [c for c in df_out.columns if c.endswith("_BaseWeight")]

    front_cols  = ["Region", "Size"]
    front_cols  = [c for c in front_cols if c in df_out.columns]

    df_out = df_out[front_cols + sample_cols + bw_cols]

    # Build two separate small dataframes: one for sample, one for base weight
    id_cols       = [c for c in ["Region", "Size"] if c in df_out.columns]
    df_samples    = df_out[id_cols + sample_cols]
    df_baseweight = df_out[id_cols + bw_cols]

    # Add a "SampleCellTotal" column to df_baseweight
    df_baseweight["SampleCellTotal"] = df_samples[sample_cols].sum(axis=1)

    # round baseweight columns
    subset_bw_cols = [c for c in df_baseweight.columns if c.endswith("_BaseWeight")]
    df_baseweight[subset_bw_cols] = df_baseweight[subset_bw_cols].round(1)

    # figure out color-scale range
    norm_bw_cols = [c for c in subset_bw_cols if c != "GrandTotal_BaseWeight"]
    if len(df_out) > 1 and norm_bw_cols:
        norm_df      = df_out.iloc[:-1]
        global_min   = norm_df[norm_bw_cols].min().min()
        global_max   = norm_df[norm_bw_cols].max().max()
        global_mid   = np.percentile(norm_df[norm_bw_cols].stack(), 50)
    else:
        global_min = global_mid = global_max = 0

    output = io.BytesIO()
    pivot_propsample_rounded = pivot_propsample.round(0).astype(int)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1) Samples sheet
        sheet_name_samples = f"{scenario_label}Allocated Sample" if scenario_label else "Allocated Sample"
        df_samples.to_excel(writer, sheet_name=sheet_name_samples,
                            startrow=0, startcol=0, index=False)

        # 2) Proportional Sample
        sheet_name_prop = f"{scenario_label}Proportional Sample" if scenario_label else "Proportional Sample"
        pivot_propsample_rounded.reset_index().to_excel(
            writer, sheet_name=sheet_name_prop, index=False
        )
        ws_prop = writer.sheets[sheet_name_prop]
        first_row = 2
        last_row  = ws_prop.max_row
        first_col = 3
        last_col  = ws_prop.max_column
        for row in ws_prop.iter_rows(min_row=first_row, max_row=last_row,
                                     min_col=first_col,  max_col=last_col):
            for cell in row:
                cell.number_format = "0"

        # 3) Sample_with_baseweight (the combined df_out)
        sheet_name_sw = f"{scenario_label}Sample_with_baseweight" if scenario_label else "Sample_with_baseweight"
        df_out.to_excel(writer, sheet_name=sheet_name_sw,
                        startrow=0, startcol=0, index=False)
        ws = writer.sheets[sheet_name_sw]

        def make_rule():
            return ColorScaleRule(
                start_type="num", start_value=global_min, start_color="00FF00",
                mid_type="num",   mid_value=global_mid,  mid_color="FFFF00",
                end_type="num",   end_value=global_max,  end_color="FF0000",
            )

        for col_name in norm_bw_cols:
            col_idx   = df_out.columns.get_loc(col_name) + 1   # 1-based
            excel_col = get_column_letter(col_idx)
            first_row = 2
            last_row  = n_rows
            rng       = f"{excel_col}{first_row}:{excel_col}{last_row}"
            ws.conditional_formatting.add(rng, make_rule())
            for cell in ws[f"{excel_col}{first_row}":f"{excel_col}{last_row}"]:
                cell[0].number_format = "0.0"

        if "GrandTotal_BaseWeight" in df_out.columns:
            gt_idx   = df_out.columns.get_loc("GrandTotal_BaseWeight") + 1
            gt_col   = get_column_letter(gt_idx)
            for cell in ws[f"{gt_col}2":f"{gt_col}{n_rows+1}"]:
                cell[0].number_format = "0.0"

        # 4) Population
        sheet_name_pop = f"{scenario_label}Population" if scenario_label else "Population"
        pivot_population.to_excel(writer, sheet_name=sheet_name_pop)

    output.seek(0)
    return output.getvalue()

###############################################################################
# 2) PRE-CHECK
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
        if pop_i <=0:
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
                "ShortBy": lb- fmax,
                "Reason":"Cell min > feasible max"
            })

        feasible_max_by_i.append(fmax)

    sum_feas= sum(feasible_max_by_i)
    if sum_feas< total_sample:
        overall_conflicts.append({
            "TotalSample": total_sample,
            "SumFeasibleMax": sum_feas,
            "ShortBy": total_sample- sum_feas,
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

###############################################################################
# 3) SLACK DIAGNOSTIC
###############################################################################
def diagnose_infeasibility_slacks(df_long,
                                  total_sample,
                                  min_cell_size,
                                  max_cell_size,
                                  max_base_weight,
                                  dimension_mins,
                                  conversion_rate):
    n_cells= len(df_long)
    x= cp.Variable(n_cells, nonneg=True)
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
    for i in range(n_cells):
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
                "Constraint":f"CellMin i={i} (R={r},Sz={s},Ind={ind})",
                "SlackUsed": var.value,
                "Comment":"We missed the cell-level LB by this amount"
            })

    df_slack_usage= pd.DataFrame(slack_usage).sort_values("SlackUsed", ascending=False)
    return df_slack_sol, df_slack_usage, prob.status

###############################################################################
# 4) MAIN SOLVER
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
    Convert wide->long but preserve the original Region,Size,Industry ordering 
    from df_wide (which is based on the 'panel' sheet).
    """
    # meltdown
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

    # single checks
    df_overall, df_cells, df_dims= detailed_feasibility_check(
        df_long, total_sample, min_cell_size, max_cell_size, max_base_weight, conversion_rate, dimension_mins
    )
    if (not df_overall.empty) or (not df_cells.empty) or (not df_dims.empty):
        # Indicate "None" to show conflict
        return None, df_cells, df_dims, df_overall

    # do the MIP
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

    # dimension
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
        chosen_solvers= [solver_choice]+ [ss for ss in candidate_solvers if ss!= solver_choice]
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
            return row["Population"]/ row["OptimizedSample"]
        else:
            return np.nan
    df_long["BaseWeight"]= df_long.apply(basew, axis=1)

    return df_long, None, None, solver_that_succeeded

###############################################################################
# 5) ALLOCATE BETWEEN PANEL & FRESH: no fractional
###############################################################################
def allocate_panel_fresh(df_long_sol, df_panel_wide, df_fresh_wide):
    """
    If x_i is even => half each. If x_i is odd => try half+1 for panel if possible,
    else half; never exceed panelPop. remainder => fresh
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
        panelPop= panel_dict.get((reg,sz,ind),0)
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

        fresh_alloc= x_i- panel_alloc
        df_long_sol.at[i,"PanelAllocated"]= panel_alloc
        df_long_sol.at[i,"FreshAllocated"]= fresh_alloc

    return df_long_sol


def pivot_in_original_order(df_alloc, df_original_wide, col_for_alloc):
    """
    Return a wide table with same row & col order as df_original_wide,
    but fill with col_for_alloc from df_alloc.
    """
    id_cols= ["Region","Size"]
    industry_cols= [c for c in df_original_wide.columns if c not in id_cols]
    alloc_dict= {}
    for rowi in df_alloc.itertuples(index=False):
        key= (rowi.Region, rowi.Size, rowi.Industry)
        val= getattr(rowi, col_for_alloc, 0)
        alloc_dict[key]= val

    df_out= df_original_wide.copy()
    for i in range(len(df_out)):
        reg= df_out.loc[i,"Region"]
        sz= df_out.loc[i,"Size"]
        for c in industry_cols:
            df_out.at[i, c]= alloc_dict.get((reg, sz, c), 0)
    return df_out


###############################################################################
# 6) MAIN APP
###############################################################################
def main():
    title_placeholder = st.empty()
    title_placeholder.title("Survey Design")
    st.write("""
    **Features**:
    1. Two sheets: 'panel' and 'fresh'.
    2. Checkbox: if unchecked => Adjusted Universe = max(panel,fresh), if checked => sum(panel+fresh).
    3. Check single-constraint feasibility. If there's a direct conflict, show the Overall, Cell, and Dimension conflicts as tables.
    4. If still we fail at the solver => run a slack-based diagnostic to see combined conflict.
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

    # We will store dimension_mins for scenario1 and scenario2:
    dimension_mins_1= {"Region":{}, "Size":{}, "Industry":{}}
    dimension_mins_2= {"Region":{}, "Size":{}, "Industry":{}}

    if uploaded_file is not None:
        base_filename = uploaded_file.name.rsplit('.', 1)[0]
        if "_" in base_filename:
            parts = base_filename.split("_", 1)
            display_title = f"{parts[0]} for {parts[1]}"
        else:
            display_title = base_filename
        title_placeholder.title(display_title)

        try:
            df_panel_wide= pd.read_excel(uploaded_file, sheet_name="panel")
            df_fresh_wide= pd.read_excel(uploaded_file, sheet_name="fresh")
        except Exception as e:
            st.error(f"Error reading 'panel'/'fresh' => {e}")
            return

        with st.expander("Original Panel Table"):
            st.data_editor(df_panel_wide, use_container_width=True, key = "wide_panel")
        with st.expander("Original Fresh Table"):
            st.data_editor(df_fresh_wide, use_container_width=True, key = "wide_fresh")

        # build Adjusted Universe with same row & column order as panel
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
        st.data_editor(df_adjusted, use_container_width=True, key = "adjusted_universe")

        # dimension sets from df_adjusted
        all_regions= df_adjusted["Region"].dropna().unique()
        all_sizes= df_adjusted["Size"].dropna().unique()
        all_inds= industry_cols

        # Pre-compute scenario 1
        n_infinity_1= compute_n_infinity(z_score_1, margin_of_error_1, p_1)
        def sum_pop_in_dim_1(df, dim_type, val):
            subset= df[df[dim_type]== val]
            tot= 0
            for cc in all_inds:
                tot+= subset[cc].fillna(0).sum()
            return tot

        # Pre-compute scenario 2
        n_infinity_2= compute_n_infinity(z_score_2, margin_of_error_2, p_2)
        def sum_pop_in_dim_2(df, dim_type, val):
            subset= df[df[dim_type]== val]
            tot= 0
            for cc in all_inds:
                tot+= subset[cc].fillna(0).sum()
            return tot

        with st.expander("Dimension Minimum Overrides for Both Scenarios", expanded=True):
            st.write("These defaults come from the sample-size formula for each scenario separately. Override if needed.")
            st.markdown("### Scenario 1")
            st.markdown("**By Region**")
            for r in all_regions:
                pop_ = sum_pop_in_dim_1(df_adjusted,"Region", r)
                defMin= compute_fpc_min(pop_, n_infinity_1)
                user_val= st.number_input(f"Min sample for Region={r} (Scenario 1)", 
                                          min_value=0, value=int(round(defMin)), step=1, key=f"dim1_reg_{r}")
                dimension_mins_1["Region"][r]= user_val

            st.markdown("**By Size**")
            for sz in all_sizes:
                pop_ = sum_pop_in_dim_1(df_adjusted,"Size", sz)
                defMin= compute_fpc_min(pop_, n_infinity_1)
                user_val= st.number_input(f"Min sample for Size={sz} (Scenario 1)",
                                          min_value=0, value=int(round(defMin)), step=1, key=f"dim1_size_{sz}")
                dimension_mins_1["Size"][sz]= user_val

            st.markdown("**By Industry**")
            for ind_ in all_inds:
                pop_= df_adjusted[ind_].fillna(0).sum()
                defMin= compute_fpc_min(pop_, n_infinity_1)
                user_val= st.number_input(f"Min sample for Industry={ind_} (Scenario 1)",
                                          min_value=0, value=int(round(defMin)), step=1, key=f"dim1_ind_{ind_}")
                dimension_mins_1["Industry"][ind_]= user_val

            st.markdown("---")
            st.markdown("### Scenario 2")
            st.markdown("**By Region**")
            for r in all_regions:
                pop_ = sum_pop_in_dim_2(df_adjusted,"Region", r)
                defMin= compute_fpc_min(pop_, n_infinity_2)
                user_val= st.number_input(f"Min sample for Region={r} (Scenario 2)", 
                                          min_value=0, value=int(round(defMin)), step=1, key=f"dim2_reg_{r}")
                dimension_mins_2["Region"][r]= user_val

            st.markdown("**By Size**")
            for sz in all_sizes:
                pop_ = sum_pop_in_dim_2(df_adjusted,"Size", sz)
                defMin= compute_fpc_min(pop_, n_infinity_2)
                user_val= st.number_input(f"Min sample for Size={sz} (Scenario 2)",
                                          min_value=0, value=int(round(defMin)), step=1, key=f"dim2_size_{sz}")
                dimension_mins_2["Size"][sz]= user_val

            st.markdown("**By Industry**")
            for ind_ in all_inds:
                pop_= df_adjusted[ind_].fillna(0).sum()
                defMin= compute_fpc_min(pop_, n_infinity_2)
                user_val= st.number_input(f"Min sample for Industry={ind_} (Scenario 2)",
                                          min_value=0, value=int(round(defMin)), step=1, key=f"dim2_ind_{ind_}")
                dimension_mins_2["Industry"][ind_]= user_val

        if st.button("Run Optimization for Both Scenarios"):

            # We'll define a helper to do all the scenario logic so we can re-use it
            def run_scenario(label,
                             total_sample,
                             min_cell_size,
                             max_cell_size,
                             max_base_weight,
                             solver_choice,
                             dimension_mins,
                             conversion_rate,
                             z_score,
                             margin_of_error,
                             p):
                """
                Tries to run the entire pipeline for one scenario:
                1) run_optimization
                2) allocate_panel_fresh
                3) gather pivot tables
                4) return a dict of outputs
                """
                try:
                    df_long_final, df_cell_conf, df_dim_conf, solverinfo = run_optimization(
                        df_wide=df_adjusted,
                        total_sample=total_sample,
                        min_cell_size=min_cell_size,
                        max_cell_size=max_cell_size,
                        max_base_weight=max_base_weight,
                        solver_choice=solver_choice,
                        dimension_mins=dimension_mins,
                        conversion_rate=conversion_rate
                    )
                    if df_long_final is None:
                        # means single-constraint conflict
                        return {
                            "success": False,
                            "df_cell_conf": df_cell_conf,
                            "df_dim_conf": df_dim_conf,
                            "df_overall": solverinfo,
                            "solverinfo": None
                        }
                    else:
                        # solver success
                        df_alloc = allocate_panel_fresh(df_long_final, df_panel_wide, df_fresh_wide)

                        # region-wise totals
                        region_totals = (
                            df_alloc
                            .groupby("Region")[["PanelAllocated", "FreshAllocated"]]
                            .sum()
                            .assign(SampleTotal=lambda d: d["PanelAllocated"] + d["FreshAllocated"])
                            .reset_index()
                        )
                        # size-wise totals
                        size_totals = (
                            df_alloc
                            .groupby("Size")[["PanelAllocated", "FreshAllocated"]]
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

                        # reorder columns
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

                        # color scale
                        norm_bw_cols = [c for c in subset_bw_cols if c!="GrandTotal_BaseWeight"]
                        norm_df = df_combined.iloc[:-1] if len(df_combined)>1 else None
                        if norm_df is not None and not norm_df.empty and norm_bw_cols:
                            global_min = norm_df[norm_bw_cols].min().min()
                            global_max = norm_df[norm_bw_cols].max().max()
                            global_mid = np.percentile(norm_df[norm_bw_cols].stack(),50)
                        else:
                            global_min=0
                            global_mid=0
                            global_max=0
                        custom_cmap = LinearSegmentedColormap.from_list("custom", ["#00FF00","#FFFF00","#FF0000"])
                        norm_obj = TwoSlopeNorm(vmin=global_min,vcenter=global_mid,vmax=global_max)
                        def baseweight_color(v):
                            val = norm_obj(v)
                            color= mcolors.to_hex(custom_cmap(val))
                            return f'background-color: {color}'
                        def style_bwcol(series):
                            n=len(series)
                            return [baseweight_color(val) if i<n-1 else "" for i,val in enumerate(series)]

                        stcol= df_combined.style.apply(style_bwcol, subset=norm_bw_cols)

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

                        # store all results in a dictionary
                        result = {
                            "success": True,
                            "solverinfo": solverinfo,
                            "df_alloc": df_alloc,
                            "df_combined": df_combined,
                            "df_combined_style": stcol,
                            "pivot_panel": pivot_panel,
                            "pivot_fresh": pivot_fresh,
                            "region_totals": region_totals,
                            "size_totals": size_totals,
                            "pivot_pop": pivot_pop,
                            "pivot_propsample": pivot_propsample
                        }
                        return result

                except ValueError as e:
                    # no solver found feasible solution
                    err_str = str(e)
                    if "No solver found a feasible solution" in err_str:
                        # run slack-based diag
                        df_long = df_adjusted.melt(
                            id_vars=["Region","Size"],
                            var_name="Industry",
                            value_name="Population"
                        ).fillna(0)
                        diag_sol, diag_usage, diag_status = diagnose_infeasibility_slacks(
                            df_long,
                            total_sample,
                            min_cell_size,
                            max_cell_size,
                            max_base_weight,
                            dimension_mins,
                            conversion_rate
                        )
                        return {
                            "success": False,
                            "error_msg": err_str,
                            "diag_sol": diag_sol,
                            "diag_usage": diag_usage,
                            "diag_status": diag_status
                        }
                    else:
                        return {
                            "success": False,
                            "error_msg": err_str
                        }
                except Exception as e2:
                    return {
                        "success": False,
                        "error_msg": f"Solver Error: {e2}"
                    }

            # --- run scenario 1
            scenario1_result = run_scenario(
                label="(Scenario1) ",
                total_sample=total_sample_1,
                min_cell_size=min_cell_size_1,
                max_cell_size=max_cell_size_1,
                max_base_weight=max_base_weight_1,
                solver_choice=solver_choice_1,
                dimension_mins=dimension_mins_1,
                conversion_rate=conversion_rate_1,
                z_score=z_score_1,
                margin_of_error=margin_of_error_1,
                p=p_1
            )

            # --- run scenario 2
            scenario2_result = run_scenario(
                label="(Scenario2) ",
                total_sample=total_sample_2,
                min_cell_size=min_cell_size_2,
                max_cell_size=max_cell_size_2,
                max_base_weight=max_base_weight_2,
                solver_choice=solver_choice_2,
                dimension_mins=dimension_mins_2,
                conversion_rate=conversion_rate_2,
                z_score=z_score_2,
                margin_of_error=margin_of_error_2,
                p=p_2
            )

            # --------------------------- Display results for scenario 1 ---------------------------
            st.header("Scenario 1 Results")
            if not scenario1_result["success"]:
                # conflict or solver error
                if "df_cell_conf" in scenario1_result:
                    # single-constraint conflict
                    df_overall = scenario1_result["df_overall"]
                    df_dim_conf = scenario1_result["df_dim_conf"]
                    df_cell_conf = scenario1_result["df_cell_conf"]
                    st.error("Single-constraint conflict(s) found for Scenario 1.")
                    if not df_overall.empty:
                        st.subheader("Overall Conflicts (Scenario 1)")
                        st.data_editor(df_overall, use_container_width=True)
                    if not df_dim_conf.empty:
                        st.subheader("Dimension Conflicts (Scenario 1)")
                        st.data_editor(df_dim_conf, use_container_width=True)
                    if not df_cell_conf.empty:
                        st.subheader("Cell Conflicts (Scenario 1)")
                        st.data_editor(df_cell_conf, use_container_width=True)
                elif "error_msg" in scenario1_result:
                    st.error(scenario1_result["error_msg"])
                    if "diag_sol" in scenario1_result:
                        st.warning("Slack-based diagnostic output for Scenario 1:")
                        diag_sol = scenario1_result["diag_sol"]
                        diag_usage = scenario1_result.get("diag_usage", pd.DataFrame())
                        diag_status = scenario1_result.get("diag_status", "")
                        if diag_sol is not None:
                            st.dataframe(diag_sol)
                        if not diag_usage.empty:
                            st.dataframe(diag_usage)
                        if diag_status:
                            st.write(f"Diagnostic solver status: {diag_status}")
            else:
                # show success info
                solverinfo = scenario1_result["solverinfo"]
                if isinstance(solverinfo, str):
                    st.success(f"Solved with {solverinfo}")
                else:
                    st.success("Solver succeeded (unknown solver)")

                pivot_panel_1 = scenario1_result["pivot_panel"]
                pivot_fresh_1 = scenario1_result["pivot_fresh"]
                st.subheader("Optimized Sample Allocation (Scenario 1)")

                st.subheader("Panel (Scenario 1)")
                st.data_editor(pivot_panel_1, use_container_width=True)

                st.subheader("Fresh (Scenario 1)")
                st.data_editor(scenario1_result["pivot_fresh"], use_container_width=True)

                st.subheader("Allocated Sample & Base Weights (Scenario 1)")
                st.dataframe(scenario1_result["df_combined_style"])

                st.subheader("Region-wise Sample Totals (Scenario 1)")
                st.data_editor(scenario1_result["region_totals"], use_container_width=True)

                st.subheader("Size-wise Sample Totals (Scenario 1)")
                st.data_editor(scenario1_result["size_totals"], use_container_width=True)

                st.subheader("Proportional Sample (Scenario 1)")
                st.dataframe(scenario1_result["pivot_propsample"])

            # --------------------------- Display results for scenario 2 ---------------------------
            st.header("Scenario 2 Results")
            if not scenario2_result["success"]:
                # conflict or solver error
                if "df_cell_conf" in scenario2_result:
                    # single-constraint conflict
                    df_overall_2 = scenario2_result["df_overall"]
                    df_dim_conf_2 = scenario2_result["df_dim_conf"]
                    df_cell_conf_2 = scenario2_result["df_cell_conf"]
                    st.error("Single-constraint conflict(s) found for Scenario 2.")
                    if not df_overall_2.empty:
                        st.subheader("Overall Conflicts (Scenario 2)")
                        st.data_editor(df_overall_2, use_container_width=True)
                    if not df_dim_conf_2.empty:
                        st.subheader("Dimension Conflicts (Scenario 2)")
                        st.data_editor(df_dim_conf_2, use_container_width=True)
                    if not df_cell_conf_2.empty:
                        st.subheader("Cell Conflicts (Scenario 2)")
                        st.data_editor(df_cell_conf_2, use_container_width=True)
                elif "error_msg" in scenario2_result:
                    st.error(scenario2_result["error_msg"])
                    if "diag_sol" in scenario2_result:
                        st.warning("Slack-based diagnostic output for Scenario 2:")
                        diag_sol_2 = scenario2_result["diag_sol"]
                        diag_usage_2 = scenario2_result.get("diag_usage", pd.DataFrame())
                        diag_status_2 = scenario2_result.get("diag_status", "")
                        if diag_sol_2 is not None:
                            st.dataframe(diag_sol_2)
                        if not diag_usage_2.empty:
                            st.dataframe(diag_usage_2)
                        if diag_status_2:
                            st.write(f"Diagnostic solver status: {diag_status_2}")
            else:
                # show success info
                solverinfo_2 = scenario2_result["solverinfo"]
                if isinstance(solverinfo_2, str):
                    st.success(f"Solved with {solverinfo_2}")
                else:
                    st.success("Solver succeeded (unknown solver)")

                st.subheader("Optimized Sample Allocation (Scenario 2)")

                st.subheader("Panel (Scenario 2)")
                st.data_editor(scenario2_result["pivot_panel"], use_container_width=True)

                st.subheader("Fresh (Scenario 2)")
                st.data_editor(scenario2_result["pivot_fresh"], use_container_width=True)

                st.subheader("Allocated Sample & Base Weights (Scenario 2)")
                st.dataframe(scenario2_result["df_combined_style"])

                st.subheader("Region-wise Sample Totals (Scenario 2)")
                st.data_editor(scenario2_result["region_totals"], use_container_width=True)

                st.subheader("Size-wise Sample Totals (Scenario 2)")
                st.data_editor(scenario2_result["size_totals"], use_container_width=True)

                st.subheader("Proportional Sample (Scenario 2)")
                st.dataframe(scenario2_result["pivot_propsample"])

            # --------------------------- Comparison (Scenario 1 vs. Scenario 2) ---------------------------
            st.header("Comparison: Scenario 1 minus Scenario 2")

            # Only show comparison if both scenarios succeeded
            if scenario1_result["success"] and scenario2_result["success"]:
                # We'll compare the final df_combined from each scenario
                dfc1 = scenario1_result["df_combined"].copy().reset_index(drop=True)
                dfc2 = scenario2_result["df_combined"].copy().reset_index(drop=True)

                # We'll keep common columns only
                common_cols = [c for c in dfc1.columns if c in dfc2.columns]
                dfc1 = dfc1[common_cols].copy()
                dfc2 = dfc2[common_cols].copy()

                # We identify which columns are numeric
                numeric_cols = [c for c in common_cols if pd.api.types.is_numeric_dtype(dfc1[c])]
                # Non-numeric we keep as is (likely region, size in the index-level row)
                df_diff = dfc1.copy()
                for col in numeric_cols:
                    df_diff[col] = dfc1[col] - dfc2[col]

                # We'll do a custom style for numeric columns only
                def color_diff(val):
                    if pd.isna(val):
                        return ""
                    if val > 0:
                        # green if scenario1 is bigger
                        return "background-color: #b3ffb3"
                    elif val < 0:
                        # red if scenario2 is bigger
                        return "background-color: #ffb3b3"
                    else:
                        return ""

                df_diff_style = df_diff.style.applymap(color_diff, subset=numeric_cols)
                st.dataframe(df_diff_style)

            # --------------------------- Download combined HTML & Excel for both scenarios ---------------------------
            # We'll build an HTML with sections from scenario 1, scenario 2, and difference

            scenario_sections = []
            scenario_sections.append(("Scenario 1: Allocated Sample & Base Weights", 
                                      scenario1_result["df_combined_style"] if scenario1_result["success"] else pd.DataFrame()))
            scenario_sections.append(("Scenario 2: Allocated Sample & Base Weights", 
                                      scenario2_result["df_combined_style"] if scenario2_result["success"] else pd.DataFrame()))

            if scenario1_result["success"] and scenario2_result["success"]:
                scenario_sections.append(("Scenario 1 vs. 2 Difference", df_diff_style))

            html_bytes = dfs_to_html(scenario_sections, page_title=display_title)
            html_fname = f"{base_filename}_comparison_snapshot.html"
            st.download_button(
                label="🌐 Download comparison as HTML",
                data=html_bytes,
                file_name=html_fname,
                mime="text/html"
            )

            # For Excel, we'll create a single workbook with scenario1 sheets, scenario2 sheets, plus a difference sheet
            # We'll do this by writing to in-memory multiple times, then merging.
            # But simpler is to just open one writer and do them in separate sheets.

            excel_out = io.BytesIO()
            with pd.ExcelWriter(excel_out, engine="openpyxl") as writer:
                # scenario1
                if scenario1_result["success"]:
                    bytes1 = write_excel_combined_table(
                        scenario1_result["df_combined"],
                        scenario1_result["pivot_pop"].set_index(["Region","Size"]).drop(columns="GrandTotal", errors="ignore"),
                        scenario1_result["pivot_propsample"].set_index(["Region","Size"]),
                        scenario_label="Scenario1_"
                    )
                    # load as a workbook and append to a single writer
                    wb1 = openpyxl.load_workbook(io.BytesIO(bytes1))
                    for sheetname in wb1.sheetnames:
                        ws_copy = wb1[sheetname]
                        new_sheet = writer.book.copy_worksheet(ws_copy)
                        new_sheet.title = sheetname

                # scenario2
                if scenario2_result["success"]:
                    bytes2 = write_excel_combined_table(
                        scenario2_result["df_combined"],
                        scenario2_result["pivot_pop"].set_index(["Region","Size"]).drop(columns="GrandTotal", errors="ignore"),
                        scenario2_result["pivot_propsample"].set_index(["Region","Size"]),
                        scenario_label="Scenario2_"
                    )
                    wb2 = openpyxl.load_workbook(io.BytesIO(bytes2))
                    for sheetname in wb2.sheetnames:
                        ws_copy = wb2[sheetname]
                        # if the sheetname exists already, we rename
                        new_sheet = writer.book.copy_worksheet(ws_copy)
                        new_sheet.title = sheetname

                # difference
                if scenario1_result["success"] and scenario2_result["success"]:
                    diff_sheet = writer.book.create_sheet("ScenarioDiff")
                    # We'll store df_diff as is
                    for r_idx, rowvals in enumerate([df_diff.columns.tolist()]+ df_diff.values.tolist(), start=1):
                        for c_idx, val in enumerate(rowvals, start=1):
                            diff_sheet.cell(row=r_idx, column=c_idx, value=val)

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
    main()
