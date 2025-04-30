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

def write_excel_combined_table(df_combined, pivot_population, pivot_propsample):
    """
    Writes:
      - 'Combined' (row&col totals)
      - 'Population' (row&col totals)
      - 'Proportional Sample' (no totals)
    """
    df_out = df_combined.reset_index()
    n_rows = df_out.shape[0]

    sample_cols = [c for c in df_out.columns if c.endswith("_Sample")]
    bw_cols     = [c for c in df_out.columns if c.endswith("_BaseWeight")]
    id_cols     = [c for c in ["Region", "Size"] if c in df_out.columns]

    # Reorder columns
    df_out = df_out[id_cols + sample_cols + bw_cols]

    # Evaluate base-weight color-scale
    norm_bw_cols = [c for c in bw_cols if c != "GrandTotal_BaseWeight"]
    if len(df_out) > 1 and norm_bw_cols:
        norm_df    = df_out.iloc[:-1]
        global_min = norm_df[norm_bw_cols].min().min()
        global_max = norm_df[norm_bw_cols].max().max()
        global_mid = np.percentile(norm_df[norm_bw_cols].stack(), 50)
    else:
        global_min = global_mid = global_max = 0

    output = io.BytesIO()
    pivot_propsample_rounded = pivot_propsample.round(0).astype(int)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1) Allocated Sample
        df_samples_only = df_out[id_cols + sample_cols]
        df_samples_only.to_excel(writer, sheet_name="Allocated Sample", index=False)

        # 2) Proportional Sample
        pivot_propsample_rounded.reset_index().to_excel(
            writer, sheet_name="Proportional Sample", index=False
        )
        ws_prop = writer.sheets["Proportional Sample"]
        first_row = 2
        last_row  = ws_prop.max_row
        first_col = 3
        last_col  = ws_prop.max_column
        for row in ws_prop.iter_rows(min_row=first_row, max_row=last_row,
                                     min_col=first_col,  max_col=last_col):
            for cell in row:
                cell.number_format = "0"

        # 3) Sample_with_baseweight
        sheet_name = "Sample_with_baseweight"
        df_out.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]

        def make_rule():
            return ColorScaleRule(
                start_type="num", start_value=global_min, start_color="00FF00",
                mid_type="num",   mid_value=global_mid,  mid_color="FFFF00",
                end_type="num",   end_value=global_max,  end_color="FF0000",
            )

        for col_name in norm_bw_cols:
            col_idx   = df_out.columns.get_loc(col_name) + 1
            excel_col = get_column_letter(col_idx)
            first_row = 2
            last_row  = n_rows
            rng       = f"{excel_col}{first_row}:{excel_col}{last_row}"
            ws.conditional_formatting.add(rng, make_rule())
            for cell in ws[f"{excel_col}{first_row}":f"{excel_col}{last_row}"]:
                cell[0].number_format = "0.0"

        # 4) Population
        pivot_population.to_excel(writer, sheet_name="Population")

    output.seek(0)
    return output

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
                "ShortBy": lb - fmax,
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
    """
    If run_optimization fails to find a solution, we do a slack-based approach 
    to see how far we are from feasibility.
    """
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
                constraints.append(cp.sum(x[idx_list]) + sd>= req_min)

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
    obj= cp.Minimize(cp.sum(slack_vars))
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
    Convert wide->long but preserve the row structure for Region/Size/Industry.
    Return (df_long with the solution) or None if conflict is found.
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

    # Check feasibility quickly
    df_overall, df_cells, df_dims= detailed_feasibility_check(
        df_long, total_sample, min_cell_size, max_cell_size, max_base_weight, 
        conversion_rate, dimension_mins
    )
    if (not df_overall.empty) or (not df_cells.empty) or (not df_dims.empty):
        # Return None to indicate conflict
        return None, df_cells, df_dims, df_overall

    # MIP
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

    # Dimension
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
        chosen_solvers= [solver_choice]+[sol for sol in candidate_solvers if sol!= solver_choice]
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
# 5) ALLOCATE BETWEEN PANEL & FRESH: no fractional
###############################################################################
def allocate_panel_fresh(df_long_sol, df_panel_wide, df_fresh_wide):
    """
    If x_i is even => half each. If x_i is odd => try half+1 for panel if possible,
    else half; never exceed panelPop. remainder => fresh
    """
    identifier_cols= ["Region","Size"]
    panel_inds= [c for c in df_panel_wide.columns if c not in identifier_cols]

    # Build a dictionary for PanelPop
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

    # Ensure df_long_sol is not None
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

    # Two sets of parameters for demonstration (Scenario A / Scenario B)
    st.sidebar.header("Parameters for Scenario A")
    A_total_sample   = st.sidebar.number_input("Total Sample (A)", value=1000, key="A_total_sample")
    A_min_cell_size  = st.sidebar.number_input("Min Cell Size (A)", value=4, key="A_min_cell_size")
    A_max_cell_size  = st.sidebar.number_input("Max Cell Size (A)", value=40, key="A_max_cell_size")
    A_max_base_weight= st.sidebar.number_input("Max Base Weight (A)", value=600, key="A_max_base_weight")
    A_solver_choice  = st.sidebar.selectbox("Solver (A)", ["SCIP","ECOS_BB"], index=0, key="A_solver_choice")
    A_conversion_rate= st.sidebar.number_input("Conversion Rate (A)", value=0.3, step=0.01, key="A_conversion_rate")

    use_sum_universe = st.sidebar.checkbox("Use sum(panel,fresh) instead of max(panel,fresh)", value=False)

    st.sidebar.markdown("---")

    st.sidebar.header("Parameters for Scenario B")
    B_total_sample   = st.sidebar.number_input("Total Sample (B)", value=800, key="B_total_sample")
    B_min_cell_size  = st.sidebar.number_input("Min Cell Size (B)", value=4, key="B_min_cell_size")
    B_max_cell_size  = st.sidebar.number_input("Max Cell Size (B)", value=40, key="B_max_cell_size")
    B_max_base_weight= st.sidebar.number_input("Max Base Weight (B)", value=600, key="B_max_base_weight")
    B_solver_choice  = st.sidebar.selectbox("Solver (B)", ["SCIP","ECOS_BB"], index=1, key="B_solver_choice")
    B_conversion_rate= st.sidebar.number_input("Conversion Rate (B)", value=0.3, step=0.01, key="B_conversion_rate")

    # We'll do dimension minimum dictionaries for A and B
    dimension_mins_A= {"Region":{}, "Size":{}, "Industry":{}}
    dimension_mins_B= {"Region":{}, "Size":{}, "Industry":{}}

    st.sidebar.markdown("---")

    uploaded_file= st.file_uploader("Upload Excel with 'panel','fresh'", type=["xlsx"])

    if uploaded_file is not None:
        base_filename = uploaded_file.name.rsplit('.', 1)[0]
        if "_" in base_filename:
            parts = base_filename.split("_", 1)
            display_title = f"{parts[0]} for {parts[1]}"
        else:
            display_title = base_filename
        title_placeholder.title(display_title)

        try:
            df_panel= pd.read_excel(uploaded_file, sheet_name="panel")
            df_fresh= pd.read_excel(uploaded_file, sheet_name="fresh")
        except Exception as e:
            st.error(f"Error reading 'panel'/'fresh' => {e}")
            return

        with st.expander("Original Panel Table"):
            st.data_editor(df_panel, use_container_width=True)
        with st.expander("Original Fresh Table"):
            st.data_editor(df_fresh, use_container_width=True)

        # build Adjusted Universe
        df_adjusted= df_panel.copy()
        id_cols= ["Region","Size"]
        industry_cols= [c for c in df_panel.columns if c not in id_cols]
        for c in industry_cols:
            if c in df_fresh.columns:
                if use_sum_universe:
                    df_adjusted[c] = df_panel[c].fillna(0) + df_fresh[c].fillna(0)
                else:
                    df_adjusted[c] = np.maximum(df_panel[c].fillna(0), df_fresh[c].fillna(0))

        st.subheader("Adjusted Universe Table")
        st.data_editor(df_adjusted, use_container_width=True)

        # We'll just skip dimension-based overrides for brevity or do minimal:
        # For demonstration, set them all to 0:
        all_regions= df_adjusted["Region"].dropna().unique()
        all_sizes= df_adjusted["Size"].dropna().unique()
        all_inds= industry_cols
        for r in all_regions:
            dimension_mins_A["Region"][r]= 0
            dimension_mins_B["Region"][r]= 0
        for s in all_sizes:
            dimension_mins_A["Size"][s]= 0
            dimension_mins_B["Size"][s]= 0
        for i in all_inds:
            dimension_mins_A["Industry"][i]= 0
            dimension_mins_B["Industry"][i]= 0

        st.markdown("---")

        def make_pivots(df_alloc):
            # Just a quick helper to produce pivot tables
            pivot_panel = pd.pivot_table(
                df_alloc, index=["Region","Size"], columns="Industry",
                values="PanelAllocated", aggfunc='sum', fill_value=0, margins=True, margins_name="GrandTotal"
            ).reset_index()

            pivot_fresh = pd.pivot_table(
                df_alloc, index=["Region","Size"], columns="Industry",
                values="FreshAllocated", aggfunc='sum', fill_value=0, margins=True, margins_name="GrandTotal"
            ).reset_index()

            df_combined = create_combined_table_with_totals(df_alloc)
            # reorder columns
            icols = [c for c in ["Region","Size"] if c in df_combined.columns]
            gsamp = [c for c in df_combined.columns if c=="GrandTotal_Sample"]
            sample_cols = [c for c in df_combined.columns if c.endswith("_Sample") and c not in gsamp]
            bw_cols = [c for c in df_combined.columns if c.endswith("_BaseWeight")]
            df_combined = df_combined[icols + gsamp + sample_cols + bw_cols]

            # color scale for base weights
            norm_bw_cols = [c for c in bw_cols if c!="GrandTotal_BaseWeight"]
            if len(df_combined)>1 and len(norm_bw_cols)>0:
                norm_df = df_combined.iloc[:-1]
                gmin = norm_df[norm_bw_cols].min().min()
                gmax = norm_df[norm_bw_cols].max().max()
                gmid = np.percentile(norm_df[norm_bw_cols].stack(),50)
            else:
                gmin=gmid=gmax=0

            custom_cmap = LinearSegmentedColormap.from_list("custom", ["#00FF00","#FFFF00","#FF0000"])
            norm_obj = TwoSlopeNorm(vmin=gmin,vcenter=gmid,vmax=gmax)

            def baseweight_color(v):
                val = norm_obj(v)
                return f'background-color: {mcolors.to_hex(custom_cmap(val))}'

            def style_bwcols(series):
                # skip last row
                n = len(series)
                return [baseweight_color(v) if i<n-1 else "" for i,v in enumerate(series)]

            subset_bw = [c for c in df_combined.columns if c.endswith("_BaseWeight")]
            df_combined[subset_bw] = df_combined[subset_bw].round(1)

            stcol = df_combined.style.apply(style_bwcols, subset=subset_bw)

            return pivot_panel, pivot_fresh, stcol

        if st.button("Run Optimization for Both Scenarios"):
            # SCENARIO A
            try:
                dfA_long, dfA_cell_conf, dfA_dim_conf, solverA = run_optimization(
                    df_wide=df_adjusted,
                    total_sample=A_total_sample,
                    min_cell_size=A_min_cell_size,
                    max_cell_size=A_max_cell_size,
                    max_base_weight=A_max_base_weight,
                    solver_choice=A_solver_choice,
                    dimension_mins=dimension_mins_A,
                    conversion_rate=A_conversion_rate
                )
                st.header("Scenario A Results")
                if dfA_long is None:
                    st.error("Single-constraint conflict(s) in Scenario A.")
                    if solverA is not None:
                        st.write(f"Solver Info: {solverA}")
                    if dfA_cell_conf is not None and not dfA_cell_conf.empty:
                        st.subheader("Cell Conflicts (A)")
                        st.data_editor(dfA_cell_conf)
                    if dfA_dim_conf is not None and not dfA_dim_conf.empty:
                        st.subheader("Dimension Conflicts (A)")
                        st.data_editor(dfA_dim_conf)
                else:
                    st.success(f"Solved with {solverA}")
                    dfA_alloc = allocate_panel_fresh(dfA_long, df_panel, df_fresh)
                    panelA, freshA, combA = make_pivots(dfA_alloc)
                    st.subheader("Panel Allocation (A)")
                    st.data_editor(panelA, use_container_width=True)
                    st.subheader("Fresh Allocation (A)")
                    st.data_editor(freshA, use_container_width=True)
                    st.subheader("Allocated Sample & Base Weights (A)")
                    st.dataframe(combA)

            except ValueError as e:
                # Possibly no solver found feasible
                st.error(f"Scenario A error: {e}")
                if "No solver found a feasible solution" in str(e):
                    st.warning("Slack-based diagnostic for Scenario A:")
                    # build df_long quickly
                    df_longA = df_adjusted.melt(
                        id_vars=["Region","Size"], var_name="Industry", value_name="Population"
                    ).fillna(0)
                    diag_solA, diag_usageA, diag_statusA = diagnose_infeasibility_slacks(
                        df_longA,
                        A_total_sample,
                        A_min_cell_size,
                        A_max_cell_size,
                        A_max_base_weight,
                        dimension_mins_A,
                        A_conversion_rate
                    )
                    if diag_solA is not None:
                        st.dataframe(diag_solA)
                    if diag_usageA is not None and not diag_usageA.empty:
                        st.dataframe(diag_usageA)
                    st.write(f"Slack solver status: {diag_statusA}")
            except Exception as e2:
                st.error(f"Scenario A solver error: {e2}")

            # SCENARIO B
            try:
                dfB_long, dfB_cell_conf, dfB_dim_conf, solverB = run_optimization(
                    df_wide=df_adjusted,
                    total_sample=B_total_sample,
                    min_cell_size=B_min_cell_size,
                    max_cell_size=B_max_cell_size,
                    max_base_weight=B_max_base_weight,
                    solver_choice=B_solver_choice,
                    dimension_mins=dimension_mins_B,
                    conversion_rate=B_conversion_rate
                )
                st.header("Scenario B Results")
                if dfB_long is None:
                    st.error("Single-constraint conflict(s) in Scenario B.")
                    if solverB is not None:
                        st.write(f"Solver Info: {solverB}")
                    if dfB_cell_conf is not None and not dfB_cell_conf.empty:
                        st.subheader("Cell Conflicts (B)")
                        st.data_editor(dfB_cell_conf)
                    if dfB_dim_conf is not None and not dfB_dim_conf.empty:
                        st.subheader("Dimension Conflicts (B)")
                        st.data_editor(dfB_dim_conf)
                else:
                    st.success(f"Solved with {solverB}")
                    dfB_alloc = allocate_panel_fresh(dfB_long, df_panel, df_fresh)
                    panelB, freshB, combB = make_pivots(dfB_alloc)
                    st.subheader("Panel Allocation (B)")
                    st.data_editor(panelB, use_container_width=True)
                    st.subheader("Fresh Allocation (B)")
                    st.data_editor(freshB, use_container_width=True)
                    st.subheader("Allocated Sample & Base Weights (B)")
                    st.dataframe(combB)

            except ValueError as e:
                st.error(f"Scenario B error: {e}")
                if "No solver found a feasible solution" in str(e):
                    st.warning("Slack-based diagnostic for Scenario B:")
                    df_longB = df_adjusted.melt(
                        id_vars=["Region","Size"], var_name="Industry", value_name="Population"
                    ).fillna(0)
                    diag_solB, diag_usageB, diag_statusB = diagnose_infeasibility_slacks(
                        df_longB,
                        B_total_sample,
                        B_min_cell_size,
                        B_max_cell_size,
                        B_max_base_weight,
                        dimension_mins_B,
                        B_conversion_rate
                    )
                    if diag_solB is not None:
                        st.dataframe(diag_solB)
                    if diag_usageB is not None and not diag_usageB.empty:
                        st.dataframe(diag_usageB)
                    st.write(f"Slack solver status: {diag_statusB}")
            except Exception as e2:
                st.error(f"Scenario B solver error: {e2}")
    else:
        st.warning("Please upload an Excel file first.")

if __name__=="__main__":
    main()
