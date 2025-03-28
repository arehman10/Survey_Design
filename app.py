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
    """
    Computes infinite-population sample size:
        n_infinity = (Z^2 * p*(1 - p)) / (margin_of_error^2)
    """
    return (z_score ** 2) * p * (1 - p) / (margin_of_error ** 2)

def compute_fpc_min(N, n_infinity):
    """
    Finite population correction:
        n_fpc(N) = n_infinity / (1 + n_infinity / N)
    Returns 0 if N <= 0.
    """
    if N <= 0:
        return 0
    else:
        return n_infinity / (1 + n_infinity / N)

def create_combined_table_with_totals(df_long):
    """
    Given df_long with columns:
      [Region, Size, Industry, OptimizedSample, BaseWeight, Population, PropSample]

    1) Create pivot_sample, pivot_bw with margins=True to get a "GrandTotal" row.
    2) Interleave them as  (IndustryX_Sample, IndustryX_BaseWeight, ...) 
    3) The final table will have an extra row labeled "GrandTotal".
    """

    # Pivot for sample (summing across cells) with a GrandTotal row
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

    # Pivot for base weights (mean aggregator). 
    #   “GrandTotal” row is the average of the base weights.
    pivot_bw = pd.pivot_table(
        df_long,
        index=["Region", "Size"],
        columns="Industry",
        values="BaseWeight",
        aggfunc='sum',
        margins=True,
        margins_name="GrandTotal"
    )

    # We'll merge columns for both pivots
    all_industries = set(pivot_sample.columns).union(set(pivot_bw.columns))
    combined = pd.DataFrame(index=pivot_sample.index)

    for ind in all_industries:
        # sample col
        sample_col = f"{ind}_Sample"
        if ind in pivot_sample.columns:
            combined[sample_col] = pivot_sample[ind]
        else:
            combined[sample_col] = np.nan

        # base weight col
        bw_col = f"{ind}_BaseWeight"
        if ind in pivot_bw.columns:
            combined[bw_col] = pivot_bw[ind]
        else:
            combined[bw_col] = np.nan
    # Reorder columns so that the GrandTotal columns are the third (and fourth) columns.
    # Adjust the names as needed. here we assume they are "GrandTotal_Sample" and "GrandTotal_BaseWeight"
    gt_sample = "GrandTotal_Sample"
    gt_bw = "GrandTotal_BaseWeight"
    cols = list(combined.columns)

    # Remove the GrandTotal columns from their current positions
    if gt_sample in cols:
        cols.remove(gt_sample)
    if gt_bw in cols:
        cols.remove(gt_bw)

    # Insert them at the desired positions (after the first two index columns)
    # For example, if you want the sample column first and then the base weight column:
    cols.insert(0, gt_bw)   # temporary insertion at beginning
    cols.insert(0, gt_sample)
    # Now place these columns after "Region" and "Size"
    # (Assuming "Region" and "Size" are not in combined.columns since they're index labels,
    #  so you'll want them to remain as the index and only adjust the data columns)
    # If you need to adjust the data column order, you can choose your desired positions:
    desired_order = [gt_sample, gt_bw] + [col for col in cols if col not in [gt_sample, gt_bw]]
    combined = combined[desired_order]        

    return combined


def write_excel_combined_table(df_combined,
                               pivot_population,
                               pivot_propsample):
    """
    Writes:
      - "Combined" sheet for the single combined table with a "GrandTotal" row
        (color-scale on BaseWeight columns).
      - "Population" sheet
      - "Proportional Sample" sheet

    Returns a BytesIO for download.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1) Combined sheet
        sheet_name = "Combined"
        df_combined.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=0)

        ws = writer.sheets[sheet_name]
        n_rows = df_combined.shape[0]       # includes the GrandTotal row
        n_cols = df_combined.shape[1]

        # Apply a 3-color scale to columns ending "_BaseWeight"
        color_rule = ColorScaleRule(
            start_type="min", start_color="00FF00",   # green
            mid_type="percentile", mid_value=50, mid_color="FFFF00",  # yellow
            end_type="max", end_color="FF0000"        # red
        )

        # Data region in Excel starts at row=2, col=2 in 1-based indexing
        for col_idx, col_name in enumerate(df_combined.columns, start=1):
            if col_name.endswith("_BaseWeight"):
                excel_col = get_column_letter(col_idx + 1) 
                data_first_row = 2
                data_last_row  = n_rows + 1
                range_str = f"{excel_col}{data_first_row}:{excel_col}{data_last_row}"
                ws.conditional_formatting.add(range_str, color_rule)

        # 2) Population sheet
        pivot_population.to_excel(writer, sheet_name="Population")

        # 3) Proportional Sample sheet
        pivot_propsample.to_excel(writer, sheet_name="Proportional Sample")

    output.seek(0)
    return output


###############################################################################
# 2) PRE-CHECK & SOLVER
###############################################################################

def pre_check_dimensions(df_long,
                         dimension_mins,
                         conversion_rate,
                         min_cell_size,
                         max_cell_size):
    """
    1) For each cell i, feasible_max = min(pop_i, max_cell_size, ceil(pop_i*conversion_rate)).
    2) If feasible_max < min_cell_size => skip min_cell_size for that cell (not a conflict).
    3) Sum feasible_max per dimension. If sum < dimension_min => dimension conflict.
    Returns feasible_max_by_i, df_cell_warnings, df_dim_conflicts
    """
    n_cells = len(df_long)
    feasible_max_by_i = {}
    cell_warnings = []

    for i in range(n_cells):
        pop_i = df_long.loc[i, "Population"]
        if pop_i <= 0:
            feasible_max = 0
        else:
            conv_ub = math.ceil(pop_i * conversion_rate)
            feasible_max = min(pop_i, max_cell_size, conv_ub)

        feasible_max_by_i[i] = feasible_max

        if feasible_max < min_cell_size and pop_i > 0:
            cell_warnings.append({
                "CellIndex": i,
                "Region": df_long.loc[i, "Region"],
                "Size": df_long.loc[i, "Size"],
                "Industry": df_long.loc[i, "Industry"],
                "Population": pop_i,
                "FeasibleMax": feasible_max,
                "MinCellSize": min_cell_size,
                "Note": "Skipping min_cell_size for this cell"
            })

    # Dimension check
    dimension_indices = {
        "Region": {},
        "Size": {},
        "Industry": {}
    }
    for dim_type in dimension_indices:
        for val in df_long[dim_type].unique():
            idx_list = df_long.index[df_long[dim_type] == val].tolist()
            dimension_indices[dim_type][val] = idx_list

    dim_conflicts = []
    for dim_type, val_dict in dimension_indices.items():
        for dim_val, idx_list in val_dict.items():
            req_min = dimension_mins[dim_type].get(dim_val, 0)
            if req_min <= 0:
                continue
            sum_feasible = sum(feasible_max_by_i[i] for i in idx_list)
            if sum_feasible < req_min:
                dim_conflicts.append({
                    "DimType": dim_type,
                    "DimName": dim_val,
                    "RequiredMin": req_min,
                    "SumFeasibleMax": sum_feasible,
                    "Reason": "Dimension min > sum of feasible maxima"
                })

    df_cell_warnings = pd.DataFrame(cell_warnings)
    df_dim_conflicts = pd.DataFrame(dim_conflicts)
    return feasible_max_by_i, df_cell_warnings, df_dim_conflicts


def run_optimization(df_wide,
                     total_sample,
                     min_cell_size,
                     max_cell_size,
                     max_base_weight,
                     solver_choice,
                     dimension_mins,
                     conversion_rate):
    """
    1) Melt df_wide -> df_long
    2) Pre-check dimension feasibility
    3) If conflict => no solver
    4) Otherwise, build + solve CVXPY model:
       - sum(x) = total_sample
       - skip min_cell_size if feasible_max < min_cell_size
       - x[i] <= feasible_max, x[i] <= pop_i
       - sum_x_dim >= dimension_mins
       - minimize sum_squares(x - propSample)
    5) Return:
       - df_long with (OptimizedSample, BaseWeight, Population, PropSample)
       - df_cell_warnings
       - df_dim_conflicts
    """

    # Melt to long
    identifier_cols = ["Region", "Size"]
    data_industry_cols = [c for c in df_wide.columns if c not in identifier_cols]

    df_long = df_wide.melt(
        id_vars=identifier_cols,
        value_vars=data_industry_cols,
        var_name="Industry",
        value_name="Population"
    ).reset_index(drop=True)

    df_long["Population"] = df_long["Population"].fillna(0)
    n_cells = len(df_long)

    total_pop = df_long["Population"].sum()
    if total_pop > 0:
        df_long["PropSample"] = df_long["Population"] * (total_sample / total_pop)
    else:
        df_long["PropSample"] = 0

    # Pre-check
    feasible_max_by_i, df_cell_warnings, df_dim_conflicts = pre_check_dimensions(
        df_long, dimension_mins, conversion_rate, min_cell_size, max_cell_size
    )
    # If dimension conflicts => skip solver
    if not df_dim_conflicts.empty:
        return None, df_cell_warnings, df_dim_conflicts

    # Build model
    x = cp.Variable(shape=n_cells, integer=True)
    obj_deviation = cp.sum_squares(x - df_long["PropSample"].values)
    objective = cp.Minimize(obj_deviation)

    constraints = []
    constraints.append(cp.sum(x) == total_sample)

    for i in range(n_cells):
        pop_i = df_long.loc[i, "Population"]
        fm = feasible_max_by_i[i]

        if fm >= min_cell_size and pop_i > 0:
            constraints.append(x[i] >= min_cell_size)
        else:
            constraints.append(x[i] >= 0)

        constraints.append(x[i] <= fm)
        constraints.append(x[i] <= pop_i)

        # Optionally enforce base weight:
        constraints.append(x[i] >= pop_i / max_base_weight)

    # Dimension constraints
    dimension_indices = {
        "Region": {},
        "Size": {},
        "Industry": {}
    }
    for dim_type in dimension_indices:
        unique_vals = df_long[dim_type].unique()
        for val in unique_vals:
            idx_list = df_long.index[df_long[dim_type] == val].tolist()
            dimension_indices[dim_type][val] = idx_list

    for dim_type, val_dict in dimension_mins.items():
        for dim_val, req_min in val_dict.items():
            if req_min > 0:
                idx_list = dimension_indices[dim_type].get(dim_val, [])
                if idx_list:
                    constraints.append(cp.sum(x[idx_list]) >= req_min)

    problem = cp.Problem(objective, constraints)

     ###########################################################################
    # NEW: List of all solvers
    ###########################################################################
    all_solvers = [
        "ECOS_BB", "SCIP"
    ]

    # If the user-chosen solver is in the list, we prioritize that solver first,
    # then try the others. If it's not in the list for some reason, we just try all_solvers.
    if solver_choice in all_solvers:
        chosen_solvers = [solver_choice] + [s for s in all_solvers if s != solver_choice]
    else:
        chosen_solvers = all_solvers

    solution_found = False
    last_error = None
    solver_that_succeeded = None

    for solver in chosen_solvers:
        try:
            print(f"Trying solver {solver}...")
            result = problem.solve(solver=solver)
            if problem.status not in ["infeasible", "unbounded"] and x.value is not None:
                solution_found = True
                # This stores info which solver solved the model
                solver_that_succeeded = solver
                print(f"Solved with {solver}, objective={problem.value:.2f}")
                break
        except Exception as e:
            last_error = e
            print(f"Solver {solver} error: {e}")

    if not solution_found:
        raise ValueError(
            f"No solver found a feasible solution among {chosen_solvers}. "
            f"Last error was: {last_error}"
        )

    # Retrieve solution
    x_sol = np.round(x.value).astype(int)
    df_long["OptimizedSample"] = x_sol

    # BaseWeight
    def compute_bw(row):
        if row["OptimizedSample"] > 0:
            return row["Population"] / row["OptimizedSample"]
        else:
            return np.nan
    df_long["BaseWeight"] = df_long.apply(compute_bw, axis=1)

    return df_long, df_cell_warnings, df_dim_conflicts, solver_that_succeeded


###############################################################################
# 3) STREAMLIT FRONT-END
###############################################################################
def main():
    title_placeholder = st.empty()
    title_placeholder.title("Survey Design")


    all_solvers_list = [
      "ECOS_BB", "SCIP"
    ]

    

    st.write("""
    **Instructions**  
    1. Upload an Excel file with `Sheet1` containing columns: `Region`, `Size`, plus multiple Industry columns.
    2. Set your parameters in the sidebar, including dimension mins, `min_cell_size`, `conversion_rate`, etc.
    3. Output:
       - A single combined table with `(IndustryX_Sample, IndustryX_BaseWeight)` columns **and** a "GrandTotal" row.
       - Two separate sheets for **Population** and **Proportional Sample** in the Excel download.
    4. The base weight columns get color‑scaled in both Streamlit (background_gradient) and Excel (3‑color rule).
    """)

    # --- SIDEBAR ---
    st.sidebar.header("Parameters")
    total_sample = st.sidebar.number_input("Total Sample (TOTAL_SAMPLE)", value=1000)
    min_cell_size= st.sidebar.number_input("Min Cell Size (MIN_CELL_SIZE)", value=4)
    max_cell_size= st.sidebar.number_input("Max Cell Size (MAX_CELL_SIZE)", value=40)
    max_base_weight= st.sidebar.number_input("Max Base Weight (MAX_BASE_WEIGHT)", value=600)

    solver_choice= st.sidebar.selectbox("Solver", all_solvers_list, index=all_solvers_list.index("SCIP"))

    #solver_choice= st.sidebar.selectbox("Solver", ["SCIP","GLPK_MI"], index=0)

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
        
        # Read data
        try:
            df_wide= pd.read_excel(uploaded_file, sheet_name="Sheet1")
        except Exception as e:
            st.error(f"Error reading Sheet1: {e}")
            return

        # Identify dimension categories
        all_regions= df_wide["Region"].dropna().unique()
        all_sizes  = df_wide["Size"].dropna().unique()
        data_industry_cols= [c for c in df_wide.columns if c not in ["Region","Size"]]
        all_inds= data_industry_cols

        # Precompute n_infinity
        n_infinity= compute_n_infinity(z_score, margin_of_error, p)

        # Summation helper for dimension
        def sum_pop_in_dim(df, dim_type, val):
            if dim_type in ["Region","Size"]:
                subset= df[df[dim_type]==val]
                pop_sum= 0
                for ind_col in data_industry_cols:
                    pop_sum+= subset[ind_col].fillna(0).sum()
                return pop_sum
            else:
                # industry dimension => val is column name
                return df[val].fillna(0).sum()

        with st.sidebar.expander("Dimension Minimums Override", expanded=True):
            st.write("Default minima from sample-size formula. Adjust as needed.")

            # Regions
            st.markdown("**By Region**")
            for region in all_regions:
                N= sum_pop_in_dim(df_wide, "Region", region)
                default_min= compute_fpc_min(N, n_infinity)
                default_min_rounded= int(round(default_min))
                user_val= st.number_input(
                    f"Min sample for Region={region}",
                    min_value=0,
                    value=default_min_rounded,
                    step=1
                )
                dimension_mins["Region"][region]= user_val

            # Sizes
            st.markdown("**By Size**")
            for sz in all_sizes:
                N= sum_pop_in_dim(df_wide, "Size", sz)
                default_min= compute_fpc_min(N, n_infinity)
                default_min_rounded= int(round(default_min))
                user_val= st.number_input(
                    f"Min sample for Size={sz}",
                    min_value=0,
                    value=default_min_rounded,
                    step=1
                )
                dimension_mins["Size"][sz]= user_val

            # Industries
            st.markdown("**By Industry**")
            for ind in all_inds:
                N= sum_pop_in_dim(df_wide, "Industry", ind)
                default_min= compute_fpc_min(N, n_infinity)
                default_min_rounded= int(round(default_min))
                user_val= st.number_input(
                    f"Min sample for Industry={ind}",
                    min_value=0,
                    value=default_min_rounded,
                    step=1
                )
                dimension_mins["Industry"][ind]= user_val

        # RUN BUTTON
        if st.button("Run Optimization"):
            try:
                df_long_final, df_cell_warnings, df_dim_conflicts, solver_that_succeeded = run_optimization(
                    df_wide       = df_wide,
                    total_sample  = total_sample,
                    min_cell_size = min_cell_size,
                    max_cell_size = max_cell_size,
                    max_base_weight= max_base_weight,
                    solver_choice = solver_choice,
                    dimension_mins= dimension_mins,
                    conversion_rate= conversion_rate
                )

                if df_long_final is None:
                    # dimension conflicts => no solver
                    st.error("Dimension-level conflict: Some dimension min > sum of feasible maxima.")
                    st.dataframe(df_dim_conflicts)
                    st.warning("Please adjust constraints.")
                else:
                    # If any cell warnings
                    if not df_cell_warnings.empty:
                        st.warning("Some cells had feasible_max < min_cell_size => skipping minCellSize there:")
                        st.dataframe(df_cell_warnings)

                    st.success(f"Optimization successful! Solved with **{solver_that_succeeded}**.")

                    # 1) Create the combined table (IndustryX_Sample, IndustryX_BaseWeight) + GrandTotal row
                    df_combined = create_combined_table_with_totals(df_long_final)

                    # 2) Also build pivot tables for Population & PropSample
                    pivot_population = df_long_final.pivot_table(
                        index=["Region","Size"],
                        columns="Industry",
                        values="Population",
                        aggfunc='sum'
                    )
                    pivot_propsample = df_long_final.pivot_table(
                        index=["Region","Size"],
                        columns="Industry",
                        values="PropSample",
                        aggfunc='mean'
                    )
                    

                    # Show combined table, color only columns ending in _BaseWeight
                    # Identify all _BaseWeight columns (for styling)
                    subset_bw_cols = [c for c in df_combined.columns if c.endswith("_BaseWeight")]

                    # Exclude "GrandTotal_BaseWeight" from normalization calculations
                    norm_bw_cols = [c for c in subset_bw_cols if c != "GrandTotal_BaseWeight"]

                    # Exclude the last row (assumed to be the GrandTotal row) from normalization calculations
                    norm_df = df_combined.iloc[:-1]

                    # Compute global min, max, and median (50th percentile) across norm_bw_cols only
                    global_min = norm_df[norm_bw_cols].min().min()
                    global_max = norm_df[norm_bw_cols].max().max()
                    global_mid = np.percentile(norm_df[norm_bw_cols].stack(), 50)

                    # Create your custom colormap (green -> yellow -> red)
                    custom_cmap = LinearSegmentedColormap.from_list("custom", ["#00FF00", "#FFFF00", "#FF0000"])

                    # Create a normalization object with the 50th percentile as the center
                    norm_obj = TwoSlopeNorm(vmin=global_min, vcenter=global_mid, vmax=global_max)

                    # Define a function to compute background color for a given value
                    def baseweight_color(val):
                        normalized = norm_obj(val)
                        color = mcolors.to_hex(custom_cmap(normalized))
                        return f'background-color: {color}'

                    # Define a function to style each base weight column,
                    # skipping the last row (i.e. no style for the last row)
                    def style_baseweight_column(s):
                        n = len(s)
                        return [baseweight_color(val) if i < n - 1 else '' for i, val in enumerate(s)]

                    # Apply the custom function to all _BaseWeight columns
                    styled = df_combined.style.apply(style_baseweight_column, subset=norm_bw_cols)

                    st.subheader("Allocated Sample & Base Weights")
                    st.dataframe(styled)
                    ####################

                    with st.expander("Population Table"):
                        st.dataframe(pivot_population)

                    with st.expander("Proportional Sample Table"):
                        st.dataframe(pivot_propsample)

                    # Write Excel
                    excel_data = write_excel_combined_table(df_combined,
                                                            pivot_population,
                                                            pivot_propsample)

                    st.download_button(
                        label="Download Excel (Combined + Pop + PropSample)",
                        data=excel_data,
                        file_name="Optimized_Combined_with_GrandTotal.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except ValueError as e:
                st.error(f"No Feasible Solution (combined conflict): {e}")
            except Exception as e:
                st.error(f"Solver Error: {e}")
    else:
        st.warning("Please upload an Excel file first.")


if __name__ == "__main__":
    main()
