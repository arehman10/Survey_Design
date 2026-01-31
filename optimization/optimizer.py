"""
Sample allocation optimization module.

This module contains the core optimization logic using CVXPY for
constrained optimization of sample allocation.
"""

import numpy as np
import pandas as pd
import cvxpy as cp
from typing import Dict, List, Optional, Tuple, Any
import logging
from datetime import datetime

from config.settings import AppConfig, SolverConfig

logger = logging.getLogger(__name__)


class OptimizationError(Exception):
    """Custom exception for optimization errors."""
    pass


class SampleOptimizer:
    """
    Optimizer for sample allocation across multiple dimensions.
    
    This class handles the optimization of sample allocation across Region,
    Size, and Industry dimensions while respecting constraints on panel
    availability, fresh sample requirements, and minimum sample sizes.
    """
    
    def __init__(self, config: Optional[AppConfig] = None):
        """
        Initialize the optimizer.
        
        Args:
            config: Application configuration (uses default if None)
        """
        self.config = config or AppConfig()
        self.solver_config = SolverConfig()
    
    def compute_n_infinity(
        self,
        z_score: float,
        margin_of_error: float,
        p: float = 0.5
    ) -> float:
        """
        Compute infinite population sample size.
        
        Args:
            z_score: Z-score for confidence level
            margin_of_error: Desired margin of error
            p: Expected proportion (default 0.5 for maximum variability)
            
        Returns:
            Required sample size for infinite population
        """
        return (z_score ** 2) * p * (1 - p) / (margin_of_error ** 2)
    
    def compute_fpc_min(self, N: int, n_infinity: float) -> float:
        """
        Compute finite population corrected minimum sample size.
        
        Args:
            N: Population size
            n_infinity: Infinite population sample size
            
        Returns:
            Minimum sample size adjusted for finite population
        """
        if N <= 0:
            return 0
        return n_infinity / (1 + n_infinity / N)
    
    def calculate_minimums(
        self,
        data: pd.DataFrame,
        config: Dict
    ) -> Dict[str, pd.DataFrame]:
        """
        Calculate minimum sample sizes for all dimensions.
        
        Args:
            data: Input data with population information
            config: Configuration dictionary
            
        Returns:
            Dictionary containing minimum sample dataframes for each dimension
        """
        # Get z-score
        confidence_level = config.get('confidence_level', 0.95)
        z_score = self.config.get_z_score(confidence_level)
        
        margin_of_error = config.get('margin_of_error', 0.05)
        n_inf = self.compute_n_infinity(z_score, margin_of_error)
        
        minimums = {}
        
        # Region minimums
        region_pop = data.groupby('Region')['Population'].sum().reset_index()
        region_pop['MinNeeded'] = region_pop['Population'].apply(
            lambda N: self.compute_fpc_min(N, n_inf)
        )
        minimums['region'] = region_pop
        
        # Size minimums
        size_pop = data.groupby('Size')['Population'].sum().reset_index()
        size_pop['MinNeeded'] = size_pop['Population'].apply(
            lambda N: self.compute_fpc_min(N, n_inf)
        )
        minimums['size'] = size_pop
        
        # Industry minimums
        industry_pop = data.groupby('Industry')['Population'].sum().reset_index()
        industry_pop['MinNeeded'] = industry_pop['Population'].apply(
            lambda N: self.compute_fpc_min(N, n_inf)
        )
        minimums['industry'] = industry_pop
        
        logger.info(f"Calculated minimums: {len(minimums)} dimensions")
        return minimums
    
    def build_optimization_problem(
        self,
        data: pd.DataFrame,
        config: Dict,
        minimums: Dict[str, pd.DataFrame]
    ) -> Tuple[cp.Problem, Dict[str, cp.Variable]]:
        """
        Build the CVXPY optimization problem.
        
        Args:
            data: Input data
            config: Configuration dictionary
            minimums: Minimum sample requirements
            
        Returns:
            Tuple of (optimization problem, variables dictionary)
        """
        n_rows = len(data)
        
        # Decision variables
        panel = cp.Variable(n_rows, integer=True, name="panel")
        fresh = cp.Variable(n_rows, integer=True, name="fresh")
        
        # Total sample per cell
        total_sample = panel + fresh
        
        # Objective: minimize deviation from target while balancing samples
        target_sample = config.get('target_sample', 1000)
        
        # Weight for different objectives
        w1 = 1.0  # Target sample weight
        w2 = 0.1  # Base weight variance minimization
        
        objective_terms = [
            w1 * cp.square(cp.sum(total_sample) - target_sample)
        ]
        
        # Add base weight variance minimization
        populations = data['Population'].values
        base_weights = cp.multiply(populations, cp.inv_pos(total_sample + 1e-6))
        objective_terms.append(w2 * cp.sum_squares(base_weights - cp.sum(base_weights) / n_rows))
        
        objective = cp.Minimize(cp.sum(objective_terms))
        
        # Constraints
        constraints = []
        
        # 1. Non-negativity
        constraints.append(panel >= 0)
        constraints.append(fresh >= 0)
        
        # 2. Panel availability
        panel_available = data['PanelAvailable'].values
        constraints.append(panel <= panel_available)
        
        # 3. Fresh/Panel ratio constraints
        min_fresh_pct = config.get('min_fresh_pct', 0.20)
        max_panel_pct = config.get('max_panel_pct', 0.80)
        tolerance = config.get('tolerance', 0.05)
        
        total_sum = cp.sum(total_sample)
        constraints.append(cp.sum(fresh) >= min_fresh_pct * (1 - tolerance) * total_sum)
        constraints.append(cp.sum(panel) <= max_panel_pct * (1 + tolerance) * total_sum)
        
        # 4. Minimum sample constraints by dimension
        if config.get('enforce_region_mins', True):
            for _, row in minimums['region'].iterrows():
                region_mask = data['Region'] == row['Region']
                region_indices = np.where(region_mask)[0]
                if len(region_indices) > 0:
                    constraints.append(
                        cp.sum(total_sample[region_indices]) >= row['MinNeeded']
                    )
        
        if config.get('enforce_size_mins', True):
            for _, row in minimums['size'].iterrows():
                size_mask = data['Size'] == row['Size']
                size_indices = np.where(size_mask)[0]
                if len(size_indices) > 0:
                    constraints.append(
                        cp.sum(total_sample[size_indices]) >= row['MinNeeded']
                    )
        
        if config.get('enforce_industry_mins', True):
            for _, row in minimums['industry'].iterrows():
                industry_mask = data['Industry'] == row['Industry']
                industry_indices = np.where(industry_mask)[0]
                if len(industry_indices) > 0:
                    constraints.append(
                        cp.sum(total_sample[industry_indices]) >= row['MinNeeded']
                    )
        
        # 5. Target sample constraint (soft)
        target_lower = target_sample * (1 - tolerance)
        target_upper = target_sample * (1 + tolerance)
        constraints.append(total_sum >= target_lower)
        constraints.append(total_sum <= target_upper)
        
        problem = cp.Problem(objective, constraints)
        variables = {
            'panel': panel,
            'fresh': fresh,
            'total_sample': total_sample
        }
        
        logger.info(f"Built optimization problem with {len(constraints)} constraints")
        return problem, variables
    
    def solve_problem(
        self,
        problem: cp.Problem,
        config: Dict
    ) -> Dict[str, Any]:
        """
        Solve the optimization problem.
        
        Args:
            problem: CVXPY problem to solve
            config: Configuration dictionary
            
        Returns:
            Dictionary with solve results and metadata
        """
        solver_name = config.get('solver', 'ECOS')
        solver_options = self.solver_config.get_solver_options(solver_name)
        
        # Override with user-specified iterations if provided
        max_iters = config.get('max_iterations', 10000)
        if solver_name == 'ECOS':
            solver_options['max_iters'] = max_iters
        elif solver_name == 'SCS':
            solver_options['max_iters'] = max_iters
        
        logger.info(f"Solving with {solver_name} solver...")
        start_time = datetime.now()
        
        try:
            problem.solve(solver=solver_name, **solver_options)
            solve_time = (datetime.now() - start_time).total_seconds()
            
            result = {
                'success': problem.status in ['optimal', 'optimal_inaccurate'],
                'status': problem.status,
                'solver': solver_name,
                'solve_time': solve_time,
                'objective_value': problem.value if problem.value is not None else float('inf'),
                'iterations': getattr(problem.solver_stats, 'num_iters', None)
            }
            
            if not result['success']:
                result['message'] = f"Solver status: {problem.status}"
                logger.warning(f"Optimization failed: {problem.status}")
            else:
                logger.info(f"Optimization successful in {solve_time:.2f}s")
            
            return result
            
        except Exception as e:
            logger.error(f"Solver error: {e}", exc_info=True)
            return {
                'success': False,
                'status': 'error',
                'solver': solver_name,
                'message': str(e),
                'solve_time': (datetime.now() - start_time).total_seconds()
            }
    
    def extract_results(
        self,
        data: pd.DataFrame,
        variables: Dict[str, cp.Variable],
        solve_result: Dict
    ) -> Dict[str, Any]:
        """
        Extract and format optimization results.
        
        Args:
            data: Input data
            variables: Dictionary of CVXPY variables
            solve_result: Results from solver
            
        Returns:
            Dictionary with formatted results
        """
        if not solve_result['success']:
            return solve_result
        
        # Extract variable values
        panel_values = np.round(variables['panel'].value).astype(int)
        fresh_values = np.round(variables['fresh'].value).astype(int)
        total_values = panel_values + fresh_values
        
        # Create results dataframe
        df_long = data.copy()
        df_long['PanelSample'] = panel_values
        df_long['FreshSample'] = fresh_values
        df_long['SampleTotal'] = total_values
        df_long['OptimizedSample'] = total_values
        
        # Calculate base weights
        df_long['BaseWeight'] = np.where(
            df_long['SampleTotal'] > 0,
            df_long['Population'] / df_long['SampleTotal'],
            0
        )
        
        # Calculate summaries
        total_sample = int(total_values.sum())
        panel_sample = int(panel_values.sum())
        fresh_sample = int(fresh_values.sum())
        
        # Create pivot tables
        pivot_panel = pd.pivot_table(
            df_long,
            index=['Region', 'Size'],
            columns='Industry',
            values='PanelSample',
            aggfunc='sum',
            fill_value=0,
            margins=True,
            margins_name='GrandTotal'
        )
        
        pivot_fresh = pd.pivot_table(
            df_long,
            index=['Region', 'Size'],
            columns='Industry',
            values='FreshSample',
            aggfunc='sum',
            fill_value=0,
            margins=True,
            margins_name='GrandTotal'
        )
        
        # Totals by dimension
        region_totals = df_long.groupby('Region').agg({
            'PanelSample': 'sum',
            'FreshSample': 'sum',
            'SampleTotal': 'sum'
        }).reset_index().rename(columns={
            'PanelSample': 'PanelAllocated',
            'FreshSample': 'FreshAllocated'
        })
        
        size_totals = df_long.groupby('Size').agg({
            'PanelSample': 'sum',
            'FreshSample': 'sum',
            'SampleTotal': 'sum'
        }).reset_index().rename(columns={
            'PanelSample': 'PanelAllocated',
            'FreshSample': 'FreshAllocated'
        })
        
        industry_totals = df_long.groupby('Industry').agg({
            'PanelSample': 'sum',
            'FreshSample': 'sum',
            'SampleTotal': 'sum'
        }).reset_index().rename(columns={
            'PanelSample': 'PanelAllocated',
            'FreshSample': 'FreshAllocated'
        })
        
        # Combined table with samples and base weights
        df_combined = self._create_combined_table(df_long)
        
        # Package results
        result = {
            **solve_result,
            'df_long': df_long,
            'pivot_panel': pivot_panel,
            'pivot_fresh': pivot_fresh,
            'df_combined': df_combined,
            'region_totals': region_totals,
            'size_totals': size_totals,
            'industry_totals': industry_totals,
            'total_sample': total_sample,
            'panel_sample': panel_sample,
            'fresh_sample': fresh_sample,
            'panel_pct': (panel_sample / total_sample * 100) if total_sample > 0 else 0,
            'fresh_pct': (fresh_sample / total_sample * 100) if total_sample > 0 else 0
        }
        
        # Add warnings if any
        warnings = []
        if panel_sample == 0:
            warnings.append("No panel samples allocated")
        if fresh_sample == 0:
            warnings.append("No fresh samples allocated")
        if any(df_long['BaseWeight'] > 100):
            warnings.append("Some base weights exceed 100")
        
        if warnings:
            result['warnings'] = warnings
        
        return result
    
    def _create_combined_table(self, df_long: pd.DataFrame) -> pd.DataFrame:
        """
        Create combined table with samples and base weights.
        
        Args:
            df_long: Long-format results dataframe
            
        Returns:
            Combined pivot table
        """
        pivot_sample = pd.pivot_table(
            df_long,
            index=['Region', 'Size'],
            columns='Industry',
            values='OptimizedSample',
            aggfunc='sum',
            fill_value=0,
            margins=True,
            margins_name='GrandTotal'
        )
        
        pivot_bw = pd.pivot_table(
            df_long,
            index=['Region', 'Size'],
            columns='Industry',
            values='BaseWeight',
            aggfunc='sum',
            fill_value=0,
            margins=True,
            margins_name='GrandTotal'
        )
        
        # Interleave sample and base weight columns
        all_industries = set(pivot_sample.columns).union(pivot_bw.columns)
        combined = pd.DataFrame(index=pivot_sample.index)
        
        for industry in sorted(all_industries):
            sample_col = f"{industry}_Sample"
            bw_col = f"{industry}_BaseWeight"
            
            if industry in pivot_sample.columns:
                combined[sample_col] = pivot_sample[industry]
            if industry in pivot_bw.columns:
                combined[bw_col] = pivot_bw[industry]
        
        return combined
    
    def optimize(self, data: pd.DataFrame, config: Dict) -> Dict[str, Any]:
        """
        Main optimization method.
        
        Args:
            data: Input data
            config: Configuration dictionary
            
        Returns:
            Dictionary with optimization results
        """
        try:
            # Validate configuration
            errors = self.config.validate_config(config)
            if errors:
                return {
                    'success': False,
                    'message': f"Configuration errors: {'; '.join(errors)}"
                }
            
            # Calculate minimums
            minimums = self.calculate_minimums(data, config)
            
            # Build problem
            problem, variables = self.build_optimization_problem(
                data, config, minimums
            )
            
            # Solve
            solve_result = self.solve_problem(problem, config)
            
            # Extract results
            if solve_result['success']:
                results = self.extract_results(data, variables, solve_result)
                results['minimums'] = minimums
                return results
            else:
                return solve_result
                
        except Exception as e:
            logger.error(f"Optimization error: {e}", exc_info=True)
            return {
                'success': False,
                'message': f"Optimization error: {str(e)}"
            }
