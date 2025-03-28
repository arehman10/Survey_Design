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
    size_order = df_long['Size'].drop_duplicates().tolist()
    df_long['Size'] = pd.Categorical(df_long['Size'], categories=size_order, ordered=True)

    pivot_sample = pd.pivot_table(
        df_long,
        index=["Region", "Size"], 
        columns="Industry",
        values="OptimizedSample",
        aggfunc='sum',
        fill_value=0,
        margins=True,
        margins_name="GrandTotal"
    )

    pivot_bw = pd.pivot_table(
        df_long,
        index=["Region", "Size"],
        columns="Industry",
        values="BaseWeight",
        aggfunc='sum',
        margins=True,
        margins_name="GrandTotal"
    )

    all_inds = set(pivot_sample.columns).union(pivot_bw.columns)
    combined = pd.DataFrame(index=pivot_sample.index)
    for ind in all_inds:
        sample_col = f"{ind}_Sample"
        if ind in pivot_sample.columns:
            combined[sample_col] = pivot_sample[ind]
        else:
            combined[sample_col] = np.nan

        bw_col = f"{ind}_BaseWeight"
        if ind in pivot_bw.columns:
            combined[bw_col] = pivot_bw[ind]
        else:
            combined[bw_col] = np.nan

    gt_sample = "GrandTotal_Sample"
    gt_bw = "GrandTotal_BaseWeight"
    cols = list(combined.columns)
    if gt_sample in cols: cols.remove(gt_sample)
    if gt_bw in cols: cols.remove(gt_bw)
    cols.insert(0, gt_bw)
    cols.insert(0, gt_sample)
    desired_order = [gt_sample, gt_bw] + [c for c in cols if c not in [gt_sample, gt_bw]]
    combined = combined[desired_order]
    return combined


def write_excel_combined_table(df_combined, pivot_population, pivot_propsample):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sheet_name = "Combined"
        df_combined.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=0)

        ws = writer.sheets[sheet_name]
        n_rows = df_combined.shape[0]
        n_cols = df_combined.shape[1]

        color_rule = ColorScaleRule(
            start_type="min", start_color="00FF00",
            mid_type="percentile", mid_value=50, mid_color="FFFF00",
            end_type="max", end_color="FF0000"
        )
        for col_idx, col_name in enumerate(df_combined.columns, start=1):
            if col_name.endswith("_BaseWeight"):
                excel_col = get_column_letter(col_idx + 1)
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
    """
    1) For each cell, compute feasible_max. If minCell>feasible_max => immediate conflict.
    2) sum(feasible_max) vs total_sample => overall_conflict
    3) dimension min vs sum(feasible_max in that dimension) => dimension_conflict

    Returns:
      df_overall (with any total-sample conflict),
      df_cells   (with cell-level conflicts),
      df_dims    (with dimension-level conflicts).
    """

    n_cells = len(df_long)
    feasible_max_by_i = []
    lower_bound_by_i  = []
    cell_conflicts = []
    dimension_conflicts = []
    overall_conflicts = []

    for i in range(n_cells):
        pop_i = df_long.loc[i,"Population"]
        if pop_i<=0:
            fmax=0
        else:
            conv_ub = math.ceil(pop_i*conversion_rate)
            fmax = min(pop_i, max_cell_size, conv_ub)
        # base weight LB
        lb_bw = pop_i / max_base_weight
        if fmax>=min_cell_size and pop_i>0:
            lb = max(lb_bw, min_cell_size, 0)
        else:
            lb = max(lb_bw, 0)
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
        lower_bound_by_i.append(lb)

    sum_feas = sum(feasible_max_by_i)
    if sum_feas< total_sample:
        overall_conflicts.append({
            "TotalSample": total_sample,
            "SumFeasibleMax": sum_feas,
            "ShortBy": total_sample - sum_feas,
            "Reason":"Overall sum(feasible max) < total_sample"
        })

    # dimension sums
    dim_indices = {"Region": {}, "Size":{}, "Industry":{}}
    for dt in dim_indices:
        for val in df_long[dt].unique():
            idx_list = df_long.index[df_long[dt]==val].tolist()
            dim_indices[dt][val]= idx_list

    for dt, val_dict in dim_indices.items():
        for val, idx_list in val_dict.items():
            req_min = dimension_mins[dt].get(val,0)
            if req_min>0:
                sub_fmax = sum(feasible_max_by_i[i] for i in idx_list)
                if sub_fmax<req_min:
                    dimension_conflicts.append({
                        "DimType": dt,
                        "DimName": val,
                        "RequiredMin": req_min,
                        "SumFeasibleMax": sub_fmax,
                        "ShortBy": req_min - sub_fmax,
                        "Reason":"Dimension min > sum feasible"
                    })

    return (
        pd.DataFrame(overall_conflicts),
        pd.DataFrame(cell_conflicts),
        pd.DataFrame(dimension_conflicts)
    )


###############################################################################
# 3) SLACK-BASED DIAGNOSTIC FOR COMBINED CONFLICT
###############################################################################
def diagnose_infeasibility_slacks(df_long,
                                  total_sample,
                                  min_cell_size,
                                  max_cell_size,
                                  max_base_weight,
                                  dimension_mins,
                                  conversion_rate):
    """
    Build a 'slack-based' model: minimize total slack across constraints.
    We use continuous x for diagnosing. We'll see which constraints are collectively
    violated and by how much (the slack).
    Returns:
      - df_slack_sol (with x_i in SlackSolution_x),
      - df_slack_usage (which constraints needed slack),
      - problem.status
    """

    n_cells = len(df_long)
    x = cp.Variable(n_cells, nonneg=True)  # relaxed, continuous

    # Slack for total sample:
    s_tot = cp.Variable(nonneg=True, name="slack_totSample")
    constraints = []
    constraints.append( cp.sum(x) + s_tot == total_sample )

    dimension_slacks = {}
    dim_indices = {"Region": {}, "Size":{}, "Industry":{}}
    for dt in dim_indices:
        for val in df_long[dt].unique():
            idx_list = df_long.index[df_long[dt]==val].tolist()
            dim_indices[dt][val]= idx_list

    for dt, val_dict in dimension_mins.items():
        for val, req_min in val_dict.items():
            if req_min>0:
                sd = cp.Variable(nonneg=True, name=f"slack_dim_{dt}_{val}")
                dimension_slacks[(dt,val)] = sd
                idx_list = dim_indices[dt][val]
                constraints.append( cp.sum(x[idx_list]) + sd >= req_min )

    cell_slacks = []
    for i in range(n_cells):
        pop_i = df_long.loc[i,"Population"]
        conv_ub = math.ceil(pop_i*conversion_rate)
        fmax = min(pop_i, max_cell_size, conv_ub)
        lb_bw = pop_i/max_base_weight
        if fmax>= min_cell_size and pop_i>0:
            lb = max(lb_bw, min_cell_size, 0)
        else:
            lb = max(lb_bw, 0)

        s_cell = cp.Variable(nonneg=True, name=f"cellSlack_{i}")
        cell_slacks.append(s_cell)
        constraints.append( x[i] + s_cell >= lb )

    slack_vars = [s_tot] + list(dimension_slacks.values()) + cell_slacks
    obj = cp.Minimize( cp.sum(slack_vars) )
    problem = cp.Problem(obj, constraints)

    try:
        result = problem.solve(solver="ECOS", verbose=False)
    except Exception as e:
        return None, None, f"Diagnostic solver error: {e}"

    if problem.status in ["infeasible","unbounded"]:
        return None, None, f"Diagnostic problem is {problem.status}."

    x_sol = x.value
    df_slack_sol = df_long.copy()
    df_slack_sol["SlackSolution_x"] = x_sol

    slack_usage = []
    if s_tot.value>1e-8:
        slack_usage.append({
            "Constraint":"TotalSample",
            "SlackUsed": s_tot.value,
            "Comment": f"Sum(x) ended up {total_sample - s_tot.value}, short by {s_tot.value}"
        })

    for (dt,val), var in dimension_slacks.items():
        if var.value>1e-8:
            slack_usage.append({
                "Constraint": f"DimensionMin {dt}={val}",
                "SlackUsed": var.value,
                "Comment": "We missed that dimension min by this amount"
            })

    for i, var in enumerate(cell_slacks):
        if var.value>1e-8:
            r = df_long.loc[i,"Region"]
            s = df_long.loc[i,"Size"]
            ind = df_long.loc[i,"Industry"]
            slack_usage.append({
                "Constraint":f"CellMin i={i} (R={r}, Sz={s}, Ind={ind})",
                "SlackUsed": var.value,
                "Comment":"We missed the cell-level LB by this amount"
            })

    df_slack_usage = pd.DataFrame(slack_usage).sort_values("SlackUsed", ascending=False)
    return df_slack_sol, df_slack_usage, problem.status


###############################################################################
# 4) MAIN SOLVER FUNCTION
###############################################################################
def run_optimization(
    df_wide,
    total_sample,
    min_cell_size,
    max_cell_size,
    max_base_weight,
    solver_choice,
    dimension_mins,
    conversion_rate
):
    # 1) Melt to df_long
    identifier_cols = ["Region","Size"]
    data_cols = [c for c in df_wide.columns if c not in identifier_cols]

    df_long = df_wide.melt(
        id_vars=identifier_cols,
        value_vars=data_cols,
        var_name="Industry",
        value_name="Population"
    ).reset_index(drop=True)

    df_long["Population"] = df_long["Population"].fillna(0)
    n_cells = len(df_long)
    total_pop = df_long["Population"].sum()
    if total_pop>0:
        df_long["PropSample"] = df_long["Population"]*(total_sample/ total_pop)
    else:
        df_long["PropSample"] = 0

    # 2) Single-constraint checks
    df_overall, df_cells, df_dims = detailed_feasibility_check(
        df_long, total_sample, min_cell_size, max_cell_size, max_base_weight, conversion_rate, dimension_mins
    )
    if (not df_overall.empty) or (not df_cells.empty) or (not df_dims.empty):
        # return them so we can display them in nice tables
        return None, df_cells, df_dims, df_overall  # skip the solver

    # 3) Build the CVXPY MIP
    x = cp.Variable(n_cells, integer=True)
    obj_deviation = cp.sum_squares(x - df_long["PropSample"])
    objective = cp.Minimize(obj_deviation)
    constraints = [cp.sum(x)== total_sample]

    for i in range(n_cells):
        pop_i = df_long.loc[i,"Population"]
        conv_ub = math.ceil(pop_i*conversion_rate)
        fmax = min(pop_i, max_cell_size, conv_ub)
        lb_bw= pop_i/max_base_weight
        if fmax>= min_cell_size and pop_i>0:
            lb = max(lb_bw, min_cell_size, 0)
        else:
            lb = max(lb_bw, 0)
        constraints.append( x[i]>= lb )
        constraints.append( x[i]<= fmax )
        constraints.append( x[i]<= pop_i )

    # dimension constraints
    dim_idx= {"Region":{}, "Size":{}, "Industry":{}}
    for dt in dim_idx:
        for val in df_long[dt].unique():
            idx_list = df_long.index[df_long[dt]== val].tolist()
            dim_idx[dt][val] = idx_list

    for dt, val_dict in dimension_mins.items():
        for val, req_min in val_dict.items():
            if req_min>0:
                idx_list= dim_idx[dt].get(val,[])
                constraints.append( cp.sum(x[idx_list])>= req_min)

    problem = cp.Problem(objective, constraints)

    # 4) Try all solvers
    all_solvers = [
        "SCIP", "ECOS_BB"
    ]
    if solver_choice in all_solvers:
        chosen_solvers = [solver_choice]+ [s for s in all_solvers if s!= solver_choice]
    else:
        chosen_solvers = all_solvers

    solution_found = False
    solver_that_succeeded = None
    last_error= None

    for s in chosen_solvers:
        try:
            result = problem.solve(solver=s)
            if problem.status not in ["infeasible","unbounded"] and x.value is not None:
                solution_found= True
                solver_that_succeeded= s
                break
        except Exception as e:
            last_error= e

    if not solution_found:
        # MIP said "No feasible solution." We'll do the slack approach to see combined conflict
        raise ValueError(
            f"No solver found a feasible solution among {chosen_solvers}. Last error was: {last_error}"
        )

    # retrieve solution
    x_sol = np.round(x.value).astype(int)
    df_long["OptimizedSample"] = x_sol
    def basew(row):
        if row["OptimizedSample"]>0:
            return row["Population"]/ row["OptimizedSample"]
        else:
            return np.nan
    df_long["BaseWeight"] = df_long.apply(basew, axis=1)
    return df_long, None, None, solver_that_succeeded


###############################################################################
# 5) STREAMLIT FRONT-END
###############################################################################
def main():
    title_placeholder = st.empty()
    title_placeholder.title("Survey Design")

   # st.title("Sampling Optimization + Slack Diagnostics")

    all_solvers_list = [
        "SCIP",  "ECOS_BB"
    ]

    st.write("""
    **Flow**:
    1. Check single-constraint feasibility. If there's a direct conflict, show 
       the Overall, Cell, and Dimension conflicts as tables.
    2. If still we fail at the solver => run a slack-based diagnostic 
       to see combined conflict.
    """)

    st.sidebar.header("Parameters")
    total_sample = st.sidebar.number_input("Total Sample", value=1000)
    min_cell_size= st.sidebar.number_input("Min Cell Size", value=4)
    max_cell_size= st.sidebar.number_input("Max Cell Size", value=40)
    max_base_weight= st.sidebar.number_input("Max Base Weight", value=600)
    solver_choice= st.sidebar.selectbox("Solver", all_solvers_list, index=0)

    conversion_rate= st.sidebar.number_input("Conversion Rate", value=0.3, min_value=0.0, max_value=1.0, step=0.01)

    st.sidebar.markdown("---")
    st.sidebar.markdown("**Sample Size Formula Inputs**")
    z_score= st.sidebar.number_input("Z-Score", value=1.644853627, format="%.9f")
    margin_of_error= st.sidebar.number_input("Margin of Error", value=0.075, format="%.3f")
    p= st.sidebar.number_input("p (Population Proportion)", value=0.5, format="%.2f")

    uploaded_file= st.file_uploader("Upload Excel File", type=["xlsx"])
    dimension_mins= {
        "Region": {},
        "Size": {},
        "Industry": {}
    }

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
            df_wide= pd.read_excel(uploaded_file, sheet_name="Sheet1")
        except:
            st.error("Cannot read Sheet1 from that file.")
            return

        # set dynamic dimension mins
        all_regions= df_wide["Region"].dropna().unique()
        all_sizes  = df_wide["Size"].dropna().unique()
        data_cols = [c for c in df_wide.columns if c not in ["Region","Size"]]
        all_inds= data_cols

        n_infinity= compute_n_infinity(z_score, margin_of_error, p)

        def sum_pop_in_dim(df, dim_type, val):
            if dim_type in ["Region","Size"]:
                subset= df[df[dim_type]== val]
                tot=0
                for c in data_cols:
                    tot+= subset[c].fillna(0).sum()
                return tot
            else:
                # industry dimension => col
                return df[val].fillna(0).sum()

        with st.sidebar.expander("Dimension Minimum Overrides", expanded=True):
            st.write("Default from sample-size formula, can override if you want.")
            st.markdown("**By Region**")
            for r in all_regions:
                pop_ = sum_pop_in_dim(df_wide,"Region", r)
                defMin = compute_fpc_min(pop_, n_infinity)
                user_val= st.number_input(f"Min sample for Region={r}", min_value=0, value=int(round(defMin)), step=1)
                dimension_mins["Region"][r]= user_val

            st.markdown("**By Size**")
            for sz in all_sizes:
                pop_ = sum_pop_in_dim(df_wide,"Size", sz)
                defMin= compute_fpc_min(pop_, n_infinity)
                user_val= st.number_input(f"Min sample for Size={sz}", min_value=0, value=int(round(defMin)), step=1)
                dimension_mins["Size"][sz]= user_val

            st.markdown("**By Industry**")
            for ind_ in all_inds:
                pop_ = sum_pop_in_dim(df_wide,"Industry", ind_)
                defMin= compute_fpc_min(pop_, n_infinity)
                user_val= st.number_input(f"Min sample for Industry={ind_}", min_value=0, value=int(round(defMin)), step=1)
                dimension_mins["Industry"][ind_]= user_val

        if st.button("Run Optimization"):
            try:
                df_long_final, df_cell_conf, df_dim_conf, solver_info = run_optimization(
                    df_wide=df_wide,
                    total_sample=total_sample,
                    min_cell_size=min_cell_size,
                    max_cell_size=max_cell_size,
                    max_base_weight=max_base_weight,
                    solver_choice=solver_choice,
                    dimension_mins=dimension_mins,
                    conversion_rate=conversion_rate
                )

                if df_long_final is None:
                    # Means single-constraint conflict(s)
                    st.error("Single-constraint conflict(s) found. See tables below.")
                    df_overall = solver_info  # we stored them in 'solver_info'? Actually we returned them in that place, so let's handle it properly

                    # Actually we returned: 
                    # return None, df_cells, df_dims, df_overall
                    # so 'df_cell_conf' => cell conflicts, 'df_dim_conf' => dimension conflicts, 'solver_info' => overall conflicts
                    df_overall = solver_info

                    if not df_overall.empty:
                        st.subheader("Overall Conflicts")
                        st.dataframe(df_overall)
                    if not df_dim_conf.empty:
                        st.subheader("Dimension Conflicts")
                        st.dataframe(df_dim_conf)
                    if not df_cell_conf.empty:
                        st.subheader("Cell Conflicts")
                        st.dataframe(df_cell_conf)

                else:
                    # We got a feasible solution
                    if isinstance(solver_info, str):
                        st.success(f"Solved with solver={solver_info}")
                    else:
                        st.success("Solved with an unknown solver?")

                    # Make pivot
                    df_combined = create_combined_table_with_totals(df_long_final)
                    pivot_pop = df_long_final.pivot_table(index=["Region","Size"], columns="Industry", values="Population", aggfunc='sum')
                    pivot_prop= df_long_final.pivot_table(index=["Region","Size"], columns="Industry", values="PropSample", aggfunc='mean')

                    st.subheader("Allocated Sample & Base Weights")
                  #  st.dataframe(df_combined)

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
                    st.subheader("Population & Proportional")
                    with st.expander("Population Table"):
                        st.dataframe(pivot_pop)

                    with st.expander("Proportional Sample Table"):
                        st.dataframe(pivot_prop)


                    excel_data= write_excel_combined_table(df_combined, pivot_pop, pivot_prop)
                    st.download_button(
                        label="Download Excel",
                        data=excel_data,
                        file_name="Optimized_Results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except ValueError as e:
                # This is the combined conflict from MIP solver
                st.error(f"{e}")
                if "No solver found a feasible solution" in str(e):
                    st.warning("Now running slack-based diagnostic for combined infeasibility...")

                    # do slack approach
                    try:
                        # Rebuild df_long
                        ident_cols= ["Region","Size"]
                        dcols= [c for c in df_wide.columns if c not in ident_cols]
                        df_long= df_wide.melt(id_vars=ident_cols, value_vars=dcols, var_name="Industry", value_name="Population").reset_index(drop=True)
                        df_long["Population"]= df_long["Population"].fillna(0)
                        diag_sol, diag_usage, diag_status = diagnose_infeasibility_slacks(
                            df_long,
                            total_sample,
                            min_cell_size,
                            max_cell_size,
                            max_base_weight,
                            dimension_mins,
                            conversion_rate
                        )
                        if diag_sol is None:
                            st.error(f"Slack diagnostic also failed => {diag_status}")
                        else:
                            st.subheader("Slack-based partial solution (continuous)")
                            st.dataframe(diag_sol)

                            if diag_usage.empty:
                                st.info("No slack needed => Possibly the integer constraints cause no solution. The continuous relaxation is feasible.")
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
