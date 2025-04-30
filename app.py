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
    sections = [(title, dataframe), ...]
    Returns bytes of a single HTML file with scrollable <div> around each table.
    """
    parts = []
    for title, obj in sections:
        parts.append(f"<h2>{html.escape(title)}</h2>")
        parts.append('<div class="scrollbox">')
        if isinstance(obj, Styler):
            parts.append(obj.to_html())           # keeps colours!
        else:
            parts.append(obj.to_html(index=False, border=0, justify="center"))
      #  parts.append(df.to_html(index=False, border=0, justify="center"))
        parts.append("</div>")
    full = HTML_TEMPLATE.format(body_html="\n".join(parts),page_title=html.escape(page_title))
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

def write_excel_combined_table(df_combined, pivot_population, pivot_propsample):
    """
    Writes:
      - 'Combined' (row&col totals)
      - 'Population' (row&col totals)
      - 'Proportional Sample' (no totals)
    """
    df_out = df_combined.reset_index()          # Region, Size become columns
    n_rows = df_out.shape[0]                    # incl. grand-total row


        # ── 1. move index to columns ─────────────────────────────────────────────
#    df_out = df_combined.reset_index()          # Region, Size become columns

       # ── 1-bis. reorder so all *_Sample columns come first, then *_BaseWeight
    sample_cols = [c for c in df_out.columns if c.endswith("_Sample")]
    bw_cols     = [c for c in df_out.columns if c.endswith("_BaseWeight")]
    
    # keep Region / Size first (they are the first two columns after reset_index)
    front_cols  = ["Region", "Size"]            # present only if they exist
    front_cols  = [c for c in front_cols if c in df_out.columns]
    
    df_out = df_out[front_cols + sample_cols + bw_cols]   # ← *** new order ***

    # ----------------------------------------------------------
    # Build the two separate tables
    id_cols       = [c for c in ["Region", "Size"] if c in df_out.columns]
    
    sample_cols   = [c for c in df_out.columns if c.endswith("_Sample")]
    bw_cols       = [c for c in df_out.columns if c.endswith("_BaseWeight")]
    
    df_samples    = df_out[id_cols + sample_cols]                   # table 1
    df_baseweight = df_out[id_cols + bw_cols]                       # table 2
    # ----------------------------------------------------------

    # ------------------------------------------------------------
    #  A) row-level total for each Region–Size cell
    df_baseweight["SampleCellTotal"] = df_samples[sample_cols].sum(axis=1)
    
    #  B) Region-wise totals  (one row per Region)
    region_totals = (
        df_samples
          .groupby("Region")[sample_cols].sum()
          .sum(axis=1)                      # collapse all industry cols
          .reset_index(name="SampleRegionTotal")
    )
    
    #  C) Size-wise totals  (one row per Size)
    size_totals = (
        df_samples
          .groupby("Size")[sample_cols].sum()
          .sum(axis=1)
          .reset_index(name="SampleSizeTotal")
    )


    
    n_rows = df_out.shape[0]
    
    subset_bw_cols = [c for c in df_out.columns if c.endswith("_BaseWeight")]
    norm_bw_cols   = [c for c in subset_bw_cols
                      if c != "GrandTotal_BaseWeight"]  # exclude total
    df_out[norm_bw_cols] = df_out[norm_bw_cols].round(1)

        # ------------------------------------------------------------------
    # all columns that really are base-weights  (incl. GrandTotal)
    bw_cols = [c for c in df_out.columns if c.endswith("_BaseWeight")]
    df_out[bw_cols] = df_out[bw_cols].round(1)


    if norm_bw_cols:                              # avoid empty slice errors
        norm_df      = df_out.iloc[:-1]           # drop grand-total row
        global_min   = norm_df[norm_bw_cols].min().min()
        global_max   = norm_df[norm_bw_cols].max().max()
        global_mid   = np.percentile(norm_df[norm_bw_cols].stack(), 50)
    else:
        global_min = global_mid = global_max = 0  # fallback


    output = io.BytesIO()
    pivot_propsample_rounded = pivot_propsample.round(0).astype(int)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_out = df_out.drop(columns=["GrandTotal_BaseWeight"], errors="ignore")

         # 1) Samples sheet – no colouring needed
        df_samples.to_excel(writer, sheet_name="Allocated Sample",
                            startrow=0, startcol=0, index=False)

        # Proportional Sample – write the rounded version
        pivot_propsample_rounded.reset_index().to_excel(
            writer, sheet_name="Proportional Sample", index=False)

         # --------------------------------------------------------------
        # force 0-decimal display in the Proportional Sample sheet
        ws_prop = writer.sheets["Proportional Sample"]
        first_row = 2                                       # data start
        last_row  = ws_prop.max_row
        first_col = 3                                       # skip Region & Size
        last_col  = ws_prop.max_column
    
        for row in ws_prop.iter_rows(min_row=first_row, max_row=last_row,
                                     min_col=first_col,  max_col=last_col):
            for cell in row:
                cell.number_format = "0"                    # 0 decimal places
        # 2) BaseWeights sheet – this is where you apply rounding,
        #    colour scale, number format, etc.
#        df_baseweight.to_excel(writer, sheet_name="BaseWeights",
 #                              startrow=0, startcol=0, index=False)
    
#        ws = writer.sheets["BaseWeights"]   # all formatting applies to this sheet

       
       # df_out = df_combined.reset_index()
        sheet_name = "Sample_with_baseweight"

         # 2) Append a blank row, then the Region totals, then Size totals
   #     start_row = df_combined.shape[0] + 2                # one blank line
   #     region_totals.to_excel(writer, sheet_name="Sample_with_baseweight",
   #                            startrow=start_row, startcol=0, index=False)
        
   #     start_row += region_totals.shape[0] + 2               # another gap
   #     size_totals.to_excel(writer, sheet_name="Sample_with_baseweight",
   #                          startrow=start_row, startcol=0, index=False)

        df_out.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=0, index=False)

        ws = writer.sheets[sheet_name]

      # Colour scale rule template
        def make_rule():
            return ColorScaleRule(
                start_type="num", start_value=global_min, start_color="00FF00",
                mid_type="num",   mid_value=global_mid,  mid_color="FFFF00",
                end_type="num",   end_value=global_max,  end_color="FF0000",
            )
                
     #   n_rows = df_combined.shape[0]
     #   n_cols = df_combined.shape[1]


        # Apply the rule to every real *_BaseWeight column
        for col_name in norm_bw_cols:
            col_idx   = df_out.columns.get_loc(col_name) + 1   # 1-based
            excel_col = get_column_letter(col_idx)
            first_row = 2                                      # data start
            last_row  = n_rows                            # inclusive
            rng       = f"{excel_col}{first_row}:{excel_col}{last_row}"
            ws.conditional_formatting.add(rng, make_rule())

             # NEW: one-decimal display format
            for cell in ws[f"{excel_col}{first_row}":f"{excel_col}{last_row}"]:
                cell[0].number_format = "0.0"
        # ------------------------------------------------------------------
        # Apply the one-decimal display format to GrandTotal_BaseWeight
        if "GrandTotal_BaseWeight" in df_out.columns:
            gt_idx   = df_out.columns.get_loc("GrandTotal_BaseWeight") + 1  # 1-based
            gt_col   = get_column_letter(gt_idx)
            for cell in ws[f"{gt_col}2":f"{gt_col}{n_rows+1}"]:
                cell[0].number_format = "0.0"        

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

# --------------------------------------------------------------------------
#  NEW  ➜  Small Utilities
# --------------------------------------------------------------------------


def sidebar_param_block(label: str, default_vals: dict, key_suffix: str):
    with st.sidebar.expander(f"{label} Parameters", expanded=(label == "Design A")):
        p = {}
        p["Total Sample"]    = st.number_input("Total Sample",  value=default_vals["Total Sample"],   key=f"sample_{key_suffix}")
        p["Min Cell Size"]   = st.number_input("Min Cell Size", value=default_vals["Min Cell Size"], key=f"mincell_{key_suffix}")
        p["Max Cell Size"]   = st.number_input("Max Cell Size", value=default_vals["Max Cell Size"], key=f"maxcell_{key_suffix}")
        p["Max Base Weight"] = st.number_input("Max Base Weight",value=default_vals["Max Base Weight"],key=f"bw_{key_suffix}")
        p["Solver"]          = st.selectbox("Solver", ["SCIP","ECOS_BB"],              key=f"solver_{key_suffix}")
        p["Conversion Rate"] = st.number_input("Conversion Rate", value=default_vals["Conversion Rate"], step=0.01, key=f"conv_{key_suffix}")
        st.markdown("---")
        st.markdown("**Sample‑size formula inputs**")
        p["Z"]   = st.number_input("Z‑score",           value=default_vals["Z"],  key=f"z_{key_suffix}")
        p["MoE"] = st.number_input("Margin of Error",   value=default_vals["MoE"], key=f"moe_{key_suffix}")
        p["p"]   = st.number_input("p (population prop.)", value=default_vals["p"],  key=f"p_{key_suffix}")
        p["UseSumUniverse"] = st.checkbox("Use sum(panel,fresh) universe", value=False, key=f"sumuni_{key_suffix}")
    return p


def dim_minimums(df_adjusted: pd.DataFrame, inds: list, params: dict, key_sfx: str):
    z, moe, p = params["Z"], params["MoE"], params["p"]
    n_inf = compute_n_infinity(z, moe, p)
    dmins = {"Region": {}, "Size": {}, "Industry": {}}
    with st.sidebar.expander(f"Dimension Minimum Overrides – {key_sfx}"):
        st.write("Defaults follow FPC formula. Override if needed.")
        st.markdown("**By Region**")
        for r in df_adjusted["Region"].dropna().unique():
            pop_r = df_adjusted.loc[df_adjusted["Region"] == r, inds].fillna(0).sum().sum()
            dmins["Region"][r] = st.number_input(f"Min sample – Region {r}", 0, value=int(round(compute_fpc_min(pop_r, n_inf))), step=1, key=f"rmin_{r}_{key_sfx}")
        st.markdown("**By Size**")
        for s in df_adjusted["Size"].dropna().unique():
            pop_s = df_adjusted.loc[df_adjusted["Size"] == s, inds].fillna(0).sum().sum()
            dmins["Size"][s] = st.number_input(f"Min sample – Size {s}", 0, value=int(round(compute_fpc_min(pop_s, n_inf))), step=1, key=f"szmin_{s}_{key_sfx}")
        st.markdown("**By Industry**")
        for ind in inds:
            pop_i = df_adjusted[ind].fillna(0).sum()
            dmins["Industry"][ind] = st.number_input(f"Min sample – Industry {ind}", 0, value=int(round(compute_fpc_min(pop_i, n_inf))), step=1, key=f"imin_{ind}_{key_sfx}")
    return dmins


# --------------------------------------------------------------------------
#  NEW  ➜  Comparison helpers
# --------------------------------------------------------------------------

def diff_tables(pivot_a: pd.DataFrame, pivot_b: pd.DataFrame, label: str):
    """Return *pivot_b – pivot_a* with styling highlighting (+/-)."""
    common_cols = [c for c in pivot_a.columns if c in pivot_b.columns and pivot_a[c].dtype != 'O']
    diff = pivot_b.copy()
    diff[common_cols] = pivot_b[common_cols] - pivot_a[common_cols]
    diff_style = diff.style.applymap(lambda v: "background-color:#daf7a6" if v>0 else ("background-color:#ffb3b3" if v<0 else ""), subset=common_cols)
    return label, diff_style


###############################################################################
# 6) MAIN APP
###############################################################################
# --------------------------------------------------------------------------
#  MAIN APP – Dual Design Version (UI tweak)
# --------------------------------------------------------------------------

def main():
    st.set_page_config(page_title="Survey Design Optimizer – Compare Designs", layout="wide")
    st.title("Survey Design Optimizer – Compare Two Designs")

    # --- Upload input file
    uploaded = st.file_uploader("Upload Excel (sheets: 'panel', 'fresh')", type=["xlsx"])
    if not uploaded:
        st.info("Please upload the Excel file to proceed.")
        return
    df_panel = pd.read_excel(uploaded, sheet_name="panel")
    df_fresh = pd.read_excel(uploaded, sheet_name="fresh")

    # --- Adjusted‑universe builder
    id_cols = ["Region", "Size"]
    industry_cols = [c for c in df_panel.columns if c not in id_cols]
    def make_universe(use_sum):
        df_ad = df_panel.copy()
        for c in industry_cols:
            if c in df_fresh.columns:
                df_ad[c] = df_panel[c].fillna(0) + df_fresh[c].fillna(0) if use_sum else np.maximum(df_panel[c].fillna(0), df_fresh[c].fillna(0))
        return df_ad

    # --- Sidebar parameters for both designs
    defaults = {"Total Sample":1000, "Min Cell Size":4, "Max Cell Size":40, "Max Base Weight":600, "Conversion Rate":0.3, "Z":1.644853627, "MoE":0.075, "p":0.5}
    pA = sidebar_param_block("Design A", defaults, "A")
    pB = sidebar_param_block("Design B", defaults, "B")

    # --- Dimension minimums
    univA, univB = make_universe(pA["UseSumUniverse"]), make_universe(pB["UseSumUniverse"])
    minsA = dim_minimums(univA, industry_cols, pA, "A")
    minsB = dim_minimums(univB, industry_cols, pB, "B")

    if not st.button("🚀 Run optimisation for both designs"):
        return

    # --- Run solver for each design
    res = {}
    for lbl, df_adj, prm, dm in [("A", univA, pA, minsA), ("B", univB, pB, minsB)]:
        df_long, *_ = run_optimization(df_adj, prm["Total Sample"], prm["Min Cell Size"], prm["Max Cell Size"], prm["Max Base Weight"], prm["Solver"], dm, prm["Conversion Rate"])
        if df_long is None:
            st.error(f"Design {lbl} infeasible – adjust minimums/params.")
            return
        res_alloc = allocate_panel_fresh(df_long, df_panel, df_fresh)
        res[lbl] = {
            "alloc": res_alloc,
            "combined": create_combined_table_with_totals(res_alloc)
        }

    # --- Tabs: A, B, Comparison
    tabA, tabB, tabC = st.tabs(["Design A", "Design B", "Comparison"])
    with tabA:
        st.dataframe(res["A"]["combined"])
    with tabB:
        st.dataframe(res["B"]["combined"])
    with tabC:
        st.header("Side‑by‑Side View")
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Design A")
            st.dataframe(res["A"]["combined"])
        with col2:
            st.subheader("Design B")
            st.dataframe(res["B"]["combined"])

    # --- Snapshot & Excel (keep diff sheet for advanced users)
    sections = [("Allocated Sample & BW (A)", res["A"]["combined"]), ("Allocated Sample & BW (B)", res["B"]["combined"])]
    html_bytes = dfs_to_html(sections, page_title="SurveyDesign_Comparison")
    st.download_button("🌐 Download HTML Snapshot", html_bytes, "survey_design_comparison.html", mime="text/html")

    wb = io.BytesIO()
    with pd.ExcelWriter(wb, engine="openpyxl") as xl:
        res["A"]["combined"].to_excel(xl, sheet_name="Combined_A")
        res["B"]["combined"].to_excel(xl, sheet_name="Combined_B")
        diff_df = res["B"]["combined"].copy()
        numcols = [c for c in diff_df.columns if diff_df[c].dtype != "O"]
        diff_df[numcols] = res["B"]["combined"][numcols] - res["A"]["combined"][numcols]
        diff_df.to_excel(xl, sheet_name="Diff_B_minus_A")
    wb.seek(0)
    st.download_button("📊 Download Excel (Both Designs)", wb, "survey_design_comparison.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    main()
