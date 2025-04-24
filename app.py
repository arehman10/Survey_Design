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

# If using GLPK:
import swiglpk as glpk

###############################################################################
# 1) HELPER FUNCTIONS
###############################################################################
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

def write_excel_combined_table(df_combined, pivot_population, pivot_propsample):
    """
    Writes:
      - 'Combined' (row&col totals)
      - 'Population' (row&col totals)
      - 'Proportional Sample' (no totals)
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        df_out = df_combined.reset_index()
        
        
        sheet_name = "Combined"
        df_out.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=0, index=False)

        ws = writer.sheets[sheet_name]
        n_rows = df_combined.shape[0]
        n_cols = df_combined.shape[1]

        color_rule = ColorScaleRule(
            start_type="min", start_color="00FF00",
            mid_type="percentile", mid_value=50, mid_color="FFFF00",
            end_type="max", end_color="FF0000"
        )
        for col_idx, col_name in enumerate(df_combined.columns, start=1):
            if col_name.endswith("_BaseWeight") and col_name != "GrandTotal_BaseWeight":
                excel_col = get_column_letter(col_idx+2)
                data_first_row = 2
                data_last_row  = n_rows + 1
                range_str = f"{excel_col}{data_first_row}:{excel_col}{data_last_row}"
                ws.conditional_formatting.add(range_str, color_rule)

        pivot_population.to_excel(writer, sheet_name="Population")
        pivot_propsample.to_excel(writer, sheet_name="Proportional Sample")

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
    # We'll store the region/size/industry order from df_wide 
    # so that after meltdown, we set them as categorical => pivot won't reorder
    # 1) row order
    region_order = df_wide["Region"].tolist()  # row by row
    size_order   = df_wide["Size"].tolist()    # row by row
    # but typically Region,Size appear the same in each row => might not be strictly unique
    # Let's do a simpler approach: store unique in the order they appear
    # We'll do a "region_order2= pd.unique(df_wide["Region"])"? But that wouldn't handle row duplicates exactly. 
    # We'll do a per-row approach -> meltdown might group them. 
    # We'll handle "Industry" ordering from columns
    # Actually simpler approach: We'll handle region & size by the unique approach, 
    # industry by the order of columns

    region_order2 = pd.unique(df_wide["Region"])
    size_order2   = pd.unique(df_wide["Size"])
    # for industry
    id_cols= ["Region","Size"]
    data_cols= [c for c in df_wide.columns if c not in id_cols]
    industry_order = data_cols[:]  # preserve order as in panel

    # meltdown
    df_long= df_wide.melt(
        id_vars= id_cols,
        value_vars=data_cols,
        var_name="Industry",
        value_name="Population"
    ).reset_index(drop=True)
    # now set the categories
    df_long["Region"] = pd.Categorical(df_long["Region"], categories=region_order2, ordered=True)
    df_long["Size"]   = pd.Categorical(df_long["Size"], categories=size_order2, ordered=True)
    df_long["Industry"]= pd.Categorical(df_long["Industry"], categories=industry_order, ordered=True)

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
        return None, df_cells, df_dims, df_overall

    # do the MIP
    n_cells= len(df_long)
    x= cp.Variable(n_cells, integer=True)
    obj_dev= cp.sum_squares(x- df_long["PropSample"])
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
    # note: these are now categorical => df_long["Region"] might be a category
    # We'll just do groupby
    for dt in dim_idx:
        cat_vals= df_long[dt].unique()
        for val in cat_vals:
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
    st.sidebar.header("Parameters")
    total_sample= st.sidebar.number_input("Total Sample", value=1000)
    min_cell_size= st.sidebar.number_input("Min Cell Size", value=4)
    max_cell_size= st.sidebar.number_input("Max Cell Size", value=40)
    max_base_weight= st.sidebar.number_input("Max Base Weight", value=600)
    solver_choice= st.sidebar.selectbox("Solver", all_solvers_list, index=0)
    conversion_rate= st.sidebar.number_input("Conversion Rate", value=0.3, step=0.01)

    use_sum_universe = st.sidebar.checkbox("Use sum(panel,fresh) instead of max(panel,fresh)", value=False)

    st.sidebar.markdown("---")
    st.sidebar.markdown("**Sample Size Formula Inputs**")
    z_score= st.sidebar.number_input("Z-Score", value=1.644853627, format="%.9f")
    margin_of_error= st.sidebar.number_input("Margin of Error", value=0.075, format="%.3f")
    p= st.sidebar.number_input("p (Population Proportion)", value=0.5, format="%.2f")

    uploaded_file= st.file_uploader("Upload Excel with 'panel','fresh'", type=["xlsx"])
    dimension_mins= {"Region":{}, "Size":{}, "Industry":{}}

    if uploaded_file is not None:
        # Update the title based on the file name
        
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
        all_inds= industry_cols  # in the panel's column order

        # sample size formula
        n_infinity= compute_n_infinity(z_score, margin_of_error, p)
        def sum_pop_in_dim(df, dim_type, val):
            subset= df[df[dim_type]== val]
            tot= 0
            for cc in all_inds:
                tot+= subset[cc].fillna(0).sum()
            return tot

        with st.sidebar.expander("Dimension Minimum Overrides", expanded=True):
            st.write("Defaults from sample-size formula, override if you like.")
            st.markdown("**By Region**")
            for r in all_regions:
                pop_ = sum_pop_in_dim(df_adjusted,"Region", r)
                defMin= compute_fpc_min(pop_, n_infinity)
                user_val= st.number_input(f"Min sample for Region={r}", min_value=0, value=int(round(defMin)), step=1)
                dimension_mins["Region"][r]= user_val

            st.markdown("**By Size**")
            for sz in all_sizes:
                pop_ = sum_pop_in_dim(df_adjusted,"Size", sz)
                defMin= compute_fpc_min(pop_, n_infinity)
                user_val= st.number_input(f"Min sample for Size={sz}", min_value=0, value=int(round(defMin)), step=1)
                dimension_mins["Size"][sz]= user_val

            st.markdown("**By Industry**")
            for ind_ in all_inds:
                pop_= df_adjusted[ind_].fillna(0).sum()
                defMin= compute_fpc_min(pop_, n_infinity)
                user_val= st.number_input(f"Min sample for Industry={ind_}", min_value=0, value=int(round(defMin)), step=1)
                dimension_mins["Industry"][ind_]= user_val

        if st.button("Run Optimization"):
            try:
                # run solver
                df_long_final, df_cell_conf, df_dim_conf, solverinfo= run_optimization(
                    df_wide= df_adjusted,
                    total_sample= total_sample,
                    min_cell_size= min_cell_size,
                    max_cell_size= max_cell_size,
                    max_base_weight= max_base_weight,
                    solver_choice= solver_choice,
                    dimension_mins= dimension_mins,
                    conversion_rate= conversion_rate
                )
                if df_long_final is None:
                    st.error("Single-constraint conflict(s).")
                    df_overall= solverinfo
                    if not df_overall.empty:
                        st.subheader("Overall Conflicts")
                        st.data_editor(df_overall, use_container_width=True)
                    if not df_dim_conf.empty:
                        st.subheader("Dimension Conflicts")
                        st.data_editor(df_dim_conf, use_container_width=True)
                    if not df_cell_conf.empty:
                        st.subheader("Cell Conflicts")
                        st.data_editor(df_cell_conf, use_container_width=True)
                else:
                    if isinstance(solverinfo, str):
                        st.success(f"Solved with {solverinfo}")
                    else:
                        st.success("Solver succeeded (unknown solver)")

                    # allocate panel/fresh
                    df_alloc= df_long_final.copy()
                    df_alloc= allocate_panel_fresh(df_alloc, df_panel_wide, df_fresh_wide)

                    # final panel & fresh pivot with row/col totals
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
                    
                    st.header("Optimized Sample Allocation")
                    st.subheader("Panel")
                    col_cfg_panel={}
                    if "Region" in pivot_panel.columns:
                        col_cfg_panel["Region"] = st.column_config.TextColumn("Region")
                    if "Size" in pivot_panel.columns:
                        col_cfg_panel["Size"] = st.column_config.TextColumn("Size")
                    st.data_editor(pivot_panel, column_config=col_cfg_panel, use_container_width=True)

                    st.subheader("Fresh")
                    col_cfg_fresh={}
                    if "Region" in pivot_fresh.columns:
                        col_cfg_fresh["Region"] = st.column_config.TextColumn("Region")
                    if "Size" in pivot_fresh.columns:
                        col_cfg_fresh["Size"] = st.column_config.TextColumn("Size")
                    st.data_editor(pivot_fresh, column_config=col_cfg_fresh, use_container_width=True)

                    # combined
                    df_combined= create_combined_table_with_totals(df_alloc)
                    df_combined_reset= df_combined.reset_index()
                    st.subheader("Allocated Sample & Base Weights")

#   Coloring happens here
                    subset_bw_cols = [c for c in df_combined.columns if c.endswith("_BaseWeight")]
                    norm_bw_cols = [c for c in subset_bw_cols if c!="GrandTotal_BaseWeight"]
                    norm_df = df_combined.iloc[:-1]
                    if not norm_df[norm_bw_cols].empty:
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
                    st.dataframe(stcol)
                    




                    col_cfg_combined={}
                    if "Region" in df_combined_reset.columns:
                        col_cfg_combined["Region"] = st.column_config.TextColumn("Region")
                    if "Size" in df_combined_reset.columns:
                        col_cfg_combined["Size"] = st.column_config.TextColumn("Size")
           #         st.data_editor(df_combined_reset, column_config=col_cfg_combined, use_container_width=True)

                    # population => row/col totals
                    pivot_pop= pd.pivot_table(
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

                    # proportional => no row/col totals
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
                    st.subheader("Proportional Sample")
                    col_cfg_prop={}
                    if "Region" in pivot_propsample.columns:
                        col_cfg_prop["Region"] = st.column_config.TextColumn("Region")
                    if "Size" in pivot_propsample.columns:
                        col_cfg_prop["Size"] = st.column_config.TextColumn("Size")
                   # st.data_editor(pivot_propsample, column_config=col_cfg_prop, use_container_width=True)
                    with st.expander("Proportional Sample Table"):
                        st.dataframe(pivot_propsample)

                    # For the final excel, we rebuild pivot_pop, pivot_propsample in un-reset form
                    pivot_pop_unreset= pd.pivot_table(
                        df_alloc,
                        index=["Region","Size"],
                        columns="Industry",
                        values="Population",
                        aggfunc='sum',
                        fill_value=0,
                        margins=True,
                        margins_name="GrandTotal",
                        sort=False
                    )
                    pivot_propsample_unreset= pd.pivot_table(
                        df_alloc,
                        index=["Region","Size"],
                        columns="Industry",
                        values="PropSample",
                        aggfunc='mean',
                        fill_value=0,
                        margins=False,
                        sort=False
                    )
                    excel_data= write_excel_combined_table(df_combined, pivot_pop_unreset, pivot_propsample_unreset)
                    st.download_button(
                        label="Download Excel",
                        data=excel_data,
                        file_name="Optimized_Results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except ValueError as e:
                st.error(str(e))
                if "No solver found a feasible solution" in str(e):
                    st.warning("Now running slack-based diagnostic for combined infeasibility...")
                    try:
                        id_cols= ["Region","Size"]
                        dcols= [cc for cc in df_adjusted.columns if cc not in id_cols]
                        df_long= df_adjusted.melt(
                            id_vars=id_cols,
                            value_vars=dcols,
                            var_name="Industry",
                            value_name="Population"
                        ).reset_index(drop=True)
                        df_long["Population"]= df_long["Population"].fillna(0)

                        diag_sol, diag_usage, diag_status= diagnose_infeasibility_slacks(
                            df_long,
                            total_sample,
                            min_cell_size,
                            max_cell_size,
                            max_base_weight,
                            dimension_mins,
                            conversion_rate
                        )
                        if diag_sol is None:
                            st.error(f"Slack diag also failed => {diag_status}")
                        else:
                            st.subheader("Slack-based partial solution (continuous relaxation)")
                            st.dataframe(diag_sol)
                            if diag_usage.empty:
                                st.info("No slack needed => Possibly integer constraints cause no solution.")
                            else:
                                st.warning("Constraints that needed slack:")
                                st.dataframe(diag_usage)
                    except Exception as e2:
                        st.error(f"Slack-based diagnostic error: {e2}")
            except Exception as e2:
                st.error(f"Solver Error: {e2}")
    else:
        st.warning("Please upload an Excel file first.")


if __name__=="__main__":
    main()
