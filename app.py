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

###############################################################################
# 1) HELPER FUNCTIONS
###############################################################################
def add_grand_total_row(df, key_col_name):
    """Return df with an extra row that holds the column totals."""
    grand = df[["PanelSample", "FreshSample", "SampleTotal"]].sum().to_frame().T
    grand[key_col_name] = "Grand Total"
    return pd.concat([df, grand], ignore_index=True)[df.columns]

def create_combined_table_with_totals(df_long):
    """
    Pivot table for (OptimizedSample, BaseWeight) with row/col totals (margins=True).
    Preserves ordering if Region/Size/Industry are categorical.
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
        s_col = f"{ind}_Sample"
        b_col = f"{ind}_BaseWeight"
        combined[s_col] = pivot_sample[ind] if ind in pivot_sample.columns else np.nan
        combined[b_col] = pivot_bw[ind] if ind in pivot_bw.columns else np.nan

    # reorder so that "GrandTotal_Sample"/"GrandTotal_BaseWeight" appear first if they exist
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

def diagnose_infeasibility_slacks(df_long,
                                  total_sample,
                                  min_cell_size,
                                  max_cell_size,
                                  max_base_weight,
                                  dimension_mins,
                                  conversion_rate):
    """
    Slack-based approach to see how far we are from feasibility if no solution is found.
    """
    import cvxpy as cp
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

def run_optimization(df_wide,
                     total_sample,
                     min_cell_size,
                     max_cell_size,
                     max_base_weight,
                     solver_choice,
                     dimension_mins,
                     conversion_rate):
    """
    Solve the integer program. Return (df_long_final, cell_conflicts, dim_conflicts, solverinfo)
    or (None, cell_conflicts, dim_conflicts, overall_conflicts) if direct conflict is found.
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

    # Feasibility check
    df_overall, df_cells, df_dims = detailed_feasibility_check(
        df_long, total_sample, min_cell_size, max_cell_size, max_base_weight, 
        conversion_rate, dimension_mins
    )
    if (not df_overall.empty) or (not df_cells.empty) or (not df_dims.empty):
        # Return None to show single-constraint conflict
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
        conv_ub= math.ceil(pop_i*conversion_rate)
        fmax= min(pop_i, max_cell_size, conv_ub)
        lb_bw= pop_i/ max_base_weight
        if fmax>= min_cell_size and pop_i>0:
            lb= max(lb_bw, min_cell_size, 0)
        else:
            lb= max(lb_bw,0)
        constraints.append(x[i]>= lb)
        constraints.append(x[i]<= fmax)
        constraints.append(x[i]<= pop_i)

    # dimension mins
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
            return row["Population"]/ row["OptimizedSample"]
        return np.nan
    df_long["BaseWeight"]= df_long.apply(basew, axis=1)

    return df_long, None, None, solver_that_succeeded

def allocate_panel_fresh(df_long_sol, df_panel_wide, df_fresh_wide):
    """
    Distribute final cell sample between PanelAllocated & FreshAllocated.
    """
    identifier_cols= ["Region","Size"]
    panel_inds= [c for c in df_panel_wide.columns if c not in identifier_cols]

    df_panel_long= df_panel_wide.melt(
        id_vars=identifier_cols,
        value_vars=panel_inds,
        var_name="Industry",
        value_name="PanelPop"
    ).reset_index(drop=True)

    panel_dict= {}
    for rowi in df_panel_long.itertuples(index=False):
        key= (rowi.Region, rowi.Size, rowi.Industry)
        val= rowi.PanelPop if pd.notna(rowi.PanelPop) else 0
        panel_dict[key]= val

    # Add columns for panel/fresh
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

###############################################################################
# 2) STREAMLIT APP
###############################################################################
def main():
    st.title("Survey Design")
    st.write("""
    **Features**:
    1. Two sheets: 'panel' and 'fresh'.
    2. Checkbox: if unchecked => Adjusted Universe = max(panel,fresh), if checked => sum(panel+fresh).
    3. Check single-constraint feasibility. If there's a direct conflict, show the Overall, Cell, and Dimension conflicts as tables.
    4. If still we fail at the solver => run a slack-based diagnostic to see combined conflict.
    """)

    # SCENARIO A inputs
    st.sidebar.header("Parameters for Scenario A")
    A_total_sample    = st.sidebar.number_input("Total Sample (A)",   value=1000)
    A_min_cell_size   = st.sidebar.number_input("Min Cell Size (A)",  value=4)
    A_max_cell_size   = st.sidebar.number_input("Max Cell Size (A)",  value=40)
    A_max_base_weight = st.sidebar.number_input("Max Base Weight (A)",value=600)
    A_solver_choice   = st.sidebar.selectbox("Solver (A)", ["SCIP","ECOS_BB"], index=0)
    A_conversion_rate = st.sidebar.number_input("Conversion Rate (A)",value=0.3, step=0.01)

    use_sum_universe  = st.sidebar.checkbox("Use sum(panel,fresh) instead of max(panel,fresh)", value=False)

    # SCENARIO B inputs
    st.sidebar.header("Parameters for Scenario B")
    B_total_sample    = st.sidebar.number_input("Total Sample (B)",   value=800)
    B_min_cell_size   = st.sidebar.number_input("Min Cell Size (B)",  value=4)
    B_max_cell_size   = st.sidebar.number_input("Max Cell Size (B)",  value=40)
    B_max_base_weight = st.sidebar.number_input("Max Base Weight (B)",value=600)
    B_solver_choice   = st.sidebar.selectbox("Solver (B)", ["SCIP","ECOS_BB"], index=1)
    B_conversion_rate = st.sidebar.number_input("Conversion Rate (B)",value=0.3, step=0.01)

    st.sidebar.markdown("---")

    # We'll let you set dimension mins from the sidebar. We'll create them after reading the file
    dimension_mins_A = {"Region":{}, "Size":{}, "Industry":{}}
    dimension_mins_B = {"Region":{}, "Size":{}, "Industry":{}}

    uploaded_file = st.file_uploader("Upload Excel with 'panel','fresh'", type=["xlsx"])
    if uploaded_file is not None:
        base_filename = uploaded_file.name.rsplit('.', 1)[0]
        try:
            df_panel= pd.read_excel(uploaded_file, sheet_name="panel")
            df_fresh= pd.read_excel(uploaded_file, sheet_name="fresh")
        except Exception as e:
            st.error(f"Error reading 'panel'/'fresh' => {e}")
            return

        with st.expander("Original Panel Table"):
            st.data_editor(df_panel, use_container_width=True, key="orig_panel")
        with st.expander("Original Fresh Table"):
            st.data_editor(df_fresh, use_container_width=True, key="orig_fresh")

        # build the Adjusted Universe
        df_adjusted = df_panel.copy()
        id_cols = ["Region","Size"]
        industry_cols = [c for c in df_panel.columns if c not in id_cols]
        for c in industry_cols:
            if c in df_fresh.columns:
                if use_sum_universe:
                    df_adjusted[c] = df_panel[c].fillna(0) + df_fresh[c].fillna(0)
                else:
                    df_adjusted[c] = np.maximum(df_panel[c].fillna(0), df_fresh[c].fillna(0))

        st.subheader("Adjusted Universe Table")
        st.data_editor(df_adjusted, use_container_width=True, key="adj_universe")

        # Gather possible dimension values
        all_regions= df_adjusted["Region"].dropna().unique()
        all_sizes  = df_adjusted["Size"].dropna().unique()
        all_inds   = industry_cols

        # Let user override dimension minimums for scenario A
        with st.sidebar.expander("Dimension Minimum Overrides (Scenario A)", expanded=False):
            for r in all_regions:
                dimension_mins_A["Region"][r] = st.number_input(f"Min sample for Region={r} (A)", min_value=0, value=0, step=1, key=f"A_reg_{r}")
            for s in all_sizes:
                dimension_mins_A["Size"][s] = st.number_input(f"Min sample for Size={s} (A)", min_value=0, value=0, step=1, key=f"A_size_{s}")
            for i in all_inds:
                dimension_mins_A["Industry"][i] = st.number_input(f"Min sample for Industry={i} (A)", min_value=0, value=0, step=1, key=f"A_ind_{i}")

        # Let user override dimension minimums for scenario B
        with st.sidebar.expander("Dimension Minimum Overrides (Scenario B)", expanded=False):
            for r in all_regions:
                dimension_mins_B["Region"][r] = st.number_input(f"Min sample for Region={r} (B)", min_value=0, value=0, step=1, key=f"B_reg_{r}")
            for s in all_sizes:
                dimension_mins_B["Size"][s] = st.number_input(f"Min sample for Size={s} (B)", min_value=0, value=0, step=1, key=f"B_size_{s}")
            for i in all_inds:
                dimension_mins_B["Industry"][i] = st.number_input(f"Min sample for Industry={i} (B)", min_value=0, value=0, step=1, key=f"B_ind_{i}")

        # Helper for pivot display
        def make_pivots(df_alloc):
            pivot_panel = pd.pivot_table(
                df_alloc, index=["Region","Size"], columns="Industry",
                values="PanelAllocated", aggfunc='sum', fill_value=0,
                margins=True, margins_name="GrandTotal"
            ).reset_index()

            pivot_fresh = pd.pivot_table(
                df_alloc, index=["Region","Size"], columns="Industry",
                values="FreshAllocated", aggfunc='sum', fill_value=0,
                margins=True, margins_name="GrandTotal"
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
                n = len(series)
                return [baseweight_color(v) if i<n-1 else "" for i,v in enumerate(series)]

            subset_bw = [c for c in df_combined.columns if c.endswith("_BaseWeight")]
            df_combined[subset_bw] = df_combined[subset_bw].round(1)
            styled = df_combined.style.apply(style_bwcols, subset=subset_bw)

            return pivot_panel, pivot_fresh, styled

        if st.button("Run Optimization for Both Scenarios"):
            # SCENARIO A
            st.header("Scenario A Results")
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
                if dfA_long is None:
                    st.error("Single-constraint conflict(s) in Scenario A.")
                    if dfA_cell_conf is not None and not dfA_cell_conf.empty:
                        st.subheader("Cell Conflicts (A)")
                        st.data_editor(dfA_cell_conf, use_container_width=True, key="A_cell_conflicts")
                    if dfA_dim_conf is not None and not dfA_dim_conf.empty:
                        st.subheader("Dimension Conflicts (A)")
                        st.data_editor(dfA_dim_conf, use_container_width=True, key="A_dim_conflicts")
                else:
                    st.success(f"Solved with solver={solverA}")
                    dfA_alloc = allocate_panel_fresh(dfA_long, df_panel, df_fresh)
                    panelA, freshA, combA = make_pivots(dfA_alloc)

                    st.subheader("Panel Allocation (A)")
                    st.data_editor(panelA, use_container_width=True, key="A_panel_alloc")
                    st.subheader("Fresh Allocation (A)")
                    st.data_editor(freshA, use_container_width=True, key="A_fresh_alloc")
                    st.subheader("Allocated Sample & Base Weights (A)")
                    st.dataframe(combA, key="A_combined")

            except ValueError as e:
                st.error(f"Scenario A error: {e}")
                if "No solver found a feasible solution" in str(e):
                    st.warning("Slack-based diagnostic for Scenario A:")
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
                        st.dataframe(diag_solA, key="A_diag_sol")
                    if diag_usageA is not None and not diag_usageA.empty:
                        st.dataframe(diag_usageA, key="A_diag_usage")
                    st.write(f"Slack solver status: {diag_statusA}")
            except Exception as e2:
                st.error(f"Scenario A solver error: {e2}")

            # SCENARIO B
            st.header("Scenario B Results")
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
                if dfB_long is None:
                    st.error("Single-constraint conflict(s) in Scenario B.")
                    if dfB_cell_conf is not None and not dfB_cell_conf.empty:
                        st.subheader("Cell Conflicts (B)")
                        st.data_editor(dfB_cell_conf, use_container_width=True, key="B_cell_conflicts")
                    if dfB_dim_conf is not None and not dfB_dim_conf.empty:
                        st.subheader("Dimension Conflicts (B)")
                        st.data_editor(dfB_dim_conf, use_container_width=True, key="B_dim_conflicts")
                else:
                    st.success(f"Solved with solver={solverB}")
                    dfB_alloc = allocate_panel_fresh(dfB_long, df_panel, df_fresh)
                    panelB, freshB, combB = make_pivots(dfB_alloc)

                    st.subheader("Panel Allocation (B)")
                    st.data_editor(panelB, use_container_width=True, key="B_panel_alloc")
                    st.subheader("Fresh Allocation (B)")
                    st.data_editor(freshB, use_container_width=True, key="B_fresh_alloc")
                    st.subheader("Allocated Sample & Base Weights (B)")
                    st.dataframe(combB, key="B_combined")

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
                        st.dataframe(diag_solB, key="B_diag_sol")
                    if diag_usageB is not None and not diag_usageB.empty:
                        st.dataframe(diag_usageB, key="B_diag_usage")
                    st.write(f"Slack solver status: {diag_statusB}")
            except Exception as e2:
                st.error(f"Scenario B solver error: {e2}")

    else:
        st.warning("Please upload an Excel file first.")


if __name__ == "__main__":
    main()
