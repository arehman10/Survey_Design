"""
Sample Allocation Optimizer - Main Application

This Streamlit application optimizes sample allocation across multiple dimensions
(Region, Size, Industry) while balancing panel and fresh samples with various constraints.

Author: Refactored Version
Date: January 2026
"""

import streamlit as st
import pandas as pd
import numpy as np
from typing import Dict, List, Tuple, Optional, Any
import logging
from pathlib import Path

# Local imports
from utils.file_handlers import FileHandler
from utils.session_manager import SessionManager
from optimization.optimizer import SampleOptimizer
from visualization.excel_generator import ExcelGenerator
from visualization.html_generator import HTMLGenerator
from config.settings import AppConfig

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class SampleAllocationApp:
    """Main application class for the Sample Allocation Optimizer."""
    
    def __init__(self):
        """Initialize the application with necessary components."""
        self.config = AppConfig()
        self.file_handler = FileHandler()
        self.session_manager = SessionManager()
        self.optimizer = SampleOptimizer()
        self.excel_generator = ExcelGenerator()
        self.html_generator = HTMLGenerator()
        
        # Initialize session state
        self._init_session_state()
    
    def _init_session_state(self):
        """Initialize Streamlit session state variables."""
        if 'uploaded_data' not in st.session_state:
            st.session_state.uploaded_data = None
        if 'optimization_results' not in st.session_state:
            st.session_state.optimization_results = {}
        if 'session_id' not in st.session_state:
            st.session_state.session_id = None
    
    def render_sidebar(self):
        """Render the sidebar with configuration options."""
        st.sidebar.title("⚙️ Configuration")
        
        # Statistical parameters
        st.sidebar.subheader("Statistical Parameters")
        confidence_level = st.sidebar.slider(
            "Confidence Level (%)",
            min_value=90,
            max_value=99,
            value=95,
            help="Confidence level for sample size calculations"
        )
        
        margin_of_error = st.sidebar.slider(
            "Margin of Error (%)",
            min_value=1,
            max_value=10,
            value=5,
            help="Desired margin of error"
        )
        
        # Optimization settings
        st.sidebar.subheader("Optimization Settings")
        solver = st.sidebar.selectbox(
            "Solver",
            options=['ECOS', 'SCS', 'GLPK_MI'],
            index=0,
            help="Optimization solver to use"
        )
        
        max_iterations = st.sidebar.number_input(
            "Max Iterations",
            min_value=1000,
            max_value=100000,
            value=10000,
            help="Maximum solver iterations"
        )
        
        return {
            'confidence_level': confidence_level,
            'margin_of_error': margin_of_error,
            'solver': solver,
            'max_iterations': max_iterations
        }
    
    def render_file_upload(self):
        """Render file upload section."""
        st.header("📁 Data Upload")
        
        uploaded_file = st.file_uploader(
            "Upload Excel file with sample data",
            type=['xlsx', 'xls'],
            help="File should contain columns: Region, Size, Industry, PanelAvailable, Population"
        )
        
        if uploaded_file is not None:
            try:
                with st.spinner("Loading data..."):
                    data = self.file_handler.load_excel(uploaded_file)
                    st.session_state.uploaded_data = data
                    st.success(f"✅ Loaded {len(data)} rows of data")
                    
                    # Show preview
                    with st.expander("📊 Data Preview"):
                        st.dataframe(data.head(10), use_container_width=True)
                        
                        # Show summary statistics
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Total Rows", len(data))
                        with col2:
                            st.metric("Unique Regions", data['Region'].nunique())
                        with col3:
                            st.metric("Unique Industries", data['Industry'].nunique())
                            
            except Exception as e:
                logger.error(f"Error loading file: {e}", exc_info=True)
                st.error(f"❌ Error loading file: {str(e)}")
                st.session_state.uploaded_data = None
        
        return st.session_state.uploaded_data
    
    def render_scenario_config(self, scenario_num: int) -> Dict[str, Any]:
        """
        Render configuration for a specific scenario.
        
        Args:
            scenario_num: Scenario number (1 or 2)
            
        Returns:
            Dictionary with scenario configuration
        """
        st.subheader(f"Scenario {scenario_num}")
        
        col1, col2 = st.columns(2)
        
        with col1:
            target_sample = st.number_input(
                f"Target Sample Size (S{scenario_num})",
                min_value=100,
                max_value=100000,
                value=1000,
                step=100,
                key=f"target_s{scenario_num}"
            )
            
            min_fresh_pct = st.slider(
                f"Min Fresh % (S{scenario_num})",
                min_value=0,
                max_value=100,
                value=20,
                key=f"min_fresh_s{scenario_num}"
            )
        
        with col2:
            max_panel_pct = st.slider(
                f"Max Panel % (S{scenario_num})",
                min_value=0,
                max_value=100,
                value=80,
                key=f"max_panel_s{scenario_num}"
            )
            
            tolerance = st.slider(
                f"Tolerance % (S{scenario_num})",
                min_value=0,
                max_value=20,
                value=5,
                key=f"tolerance_s{scenario_num}"
            )
        
        # Advanced constraints
        with st.expander(f"Advanced Constraints (S{scenario_num})"):
            enforce_industry_mins = st.checkbox(
                "Enforce industry minimums",
                value=True,
                key=f"industry_mins_s{scenario_num}"
            )
            
            enforce_region_mins = st.checkbox(
                "Enforce region minimums",
                value=True,
                key=f"region_mins_s{scenario_num}"
            )
            
            enforce_size_mins = st.checkbox(
                "Enforce size minimums",
                value=True,
                key=f"size_mins_s{scenario_num}"
            )
        
        return {
            'target_sample': target_sample,
            'min_fresh_pct': min_fresh_pct / 100,
            'max_panel_pct': max_panel_pct / 100,
            'tolerance': tolerance / 100,
            'enforce_industry_mins': enforce_industry_mins,
            'enforce_region_mins': enforce_region_mins,
            'enforce_size_mins': enforce_size_mins
        }
    
    def run_optimization(self, data: pd.DataFrame, config: Dict, scenario_configs: List[Dict]):
        """
        Run optimization for all scenarios.
        
        Args:
            data: Input data
            config: Global configuration
            scenario_configs: List of scenario-specific configurations
        """
        if data is None:
            st.warning("⚠️ Please upload data first")
            return
        
        try:
            with st.spinner("🔄 Running optimization..."):
                results = {}
                
                for i, scenario_config in enumerate(scenario_configs, 1):
                    logger.info(f"Running optimization for Scenario {i}")
                    
                    # Merge configurations
                    full_config = {**config, **scenario_config}
                    
                    # Run optimization
                    result = self.optimizer.optimize(data, full_config)
                    results[f'scenario_{i}'] = result
                    
                    if result.get('success'):
                        st.success(f"✅ Scenario {i} optimized successfully")
                    else:
                        st.error(f"❌ Scenario {i} optimization failed: {result.get('message', 'Unknown error')}")
                
                st.session_state.optimization_results = results
                return results
                
        except Exception as e:
            logger.error(f"Optimization error: {e}", exc_info=True)
            st.error(f"❌ Optimization error: {str(e)}")
            return None
    
    def render_results(self, results: Dict):
        """
        Render optimization results.
        
        Args:
            results: Dictionary containing optimization results
        """
        if not results:
            return
        
        st.header("📊 Optimization Results")
        
        # Create tabs for different views
        tabs = st.tabs(["Summary", "Detailed Results", "Comparison", "Diagnostics"])
        
        with tabs[0]:  # Summary
            self._render_summary(results)
        
        with tabs[1]:  # Detailed Results
            self._render_detailed_results(results)
        
        with tabs[2]:  # Comparison
            self._render_comparison(results)
        
        with tabs[3]:  # Diagnostics
            self._render_diagnostics(results)
    
    def _render_summary(self, results: Dict):
        """Render summary statistics."""
        st.subheader("Summary Statistics")
        
        summary_data = []
        for scenario_name, result in results.items():
            if result.get('success'):
                summary_data.append({
                    'Scenario': scenario_name.replace('scenario_', 'Scenario '),
                    'Total Sample': result.get('total_sample', 0),
                    'Panel Sample': result.get('panel_sample', 0),
                    'Fresh Sample': result.get('fresh_sample', 0),
                    'Panel %': f"{result.get('panel_pct', 0):.1f}%",
                    'Fresh %': f"{result.get('fresh_pct', 0):.1f}%",
                    'Optimization Time': f"{result.get('solve_time', 0):.2f}s"
                })
        
        if summary_data:
            st.dataframe(pd.DataFrame(summary_data), use_container_width=True)
    
    def _render_detailed_results(self, results: Dict):
        """Render detailed results for each scenario."""
        for scenario_name, result in results.items():
            if result.get('success'):
                st.subheader(scenario_name.replace('scenario_', 'Scenario '))
                
                # Show allocation tables
                if 'df_long' in result:
                    st.dataframe(result['df_long'], use_container_width=True)
    
    def _render_comparison(self, results: Dict):
        """Render comparison between scenarios."""
        if len(results) >= 2:
            st.subheader("Scenario Comparison")
            
            # Create comparison metrics
            scenarios = list(results.keys())
            if all(results[s].get('success') for s in scenarios[:2]):
                s1_result = results[scenarios[0]]
                s2_result = results[scenarios[1]]
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    diff = s1_result.get('total_sample', 0) - s2_result.get('total_sample', 0)
                    st.metric("Sample Difference", f"{diff:,.0f}")
                
                with col2:
                    diff = s1_result.get('panel_sample', 0) - s2_result.get('panel_sample', 0)
                    st.metric("Panel Difference", f"{diff:,.0f}")
                
                with col3:
                    diff = s1_result.get('fresh_sample', 0) - s2_result.get('fresh_sample', 0)
                    st.metric("Fresh Difference", f"{diff:,.0f}")
    
    def _render_diagnostics(self, results: Dict):
        """Render diagnostic information."""
        st.subheader("Optimization Diagnostics")
        
        for scenario_name, result in results.items():
            with st.expander(scenario_name.replace('scenario_', 'Scenario ')):
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write("**Status:**", result.get('status', 'Unknown'))
                    st.write("**Solver:**", result.get('solver', 'Unknown'))
                    st.write("**Solve Time:**", f"{result.get('solve_time', 0):.2f}s")
                
                with col2:
                    st.write("**Iterations:**", result.get('iterations', 'N/A'))
                    st.write("**Objective Value:**", f"{result.get('objective_value', 0):,.2f}")
                
                if 'warnings' in result and result['warnings']:
                    st.warning("⚠️ Warnings:")
                    for warning in result['warnings']:
                        st.write(f"- {warning}")
    
    def render_export_options(self, results: Dict):
        """Render export options for results."""
        if not results or not any(r.get('success') for r in results.values()):
            return
        
        st.header("📥 Export Results")
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("Generate Excel Report", use_container_width=True):
                try:
                    excel_file = self.excel_generator.generate(
                        results,
                        st.session_state.uploaded_data
                    )
                    
                    st.download_button(
                        label="📥 Download Excel Report",
                        data=excel_file,
                        file_name="sample_allocation_report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    logger.error(f"Excel generation error: {e}", exc_info=True)
                    st.error(f"❌ Error generating Excel: {str(e)}")
        
        with col2:
            if st.button("Generate HTML Report", use_container_width=True):
                try:
                    html_file = self.html_generator.generate(
                        results,
                        st.session_state.uploaded_data
                    )
                    
                    st.download_button(
                        label="📥 Download HTML Report",
                        data=html_file,
                        file_name="sample_allocation_report.html",
                        mime="text/html"
                    )
                except Exception as e:
                    logger.error(f"HTML generation error: {e}", exc_info=True)
                    st.error(f"❌ Error generating HTML: {str(e)}")
    
    def run(self):
        """Main application entry point."""
        # Page configuration
        st.set_page_config(
            page_title="Sample Allocation Optimizer",
            page_icon="📊",
            layout="wide",
            initial_sidebar_state="expanded"
        )
        
        # Header
        st.title("📊 Sample Allocation Optimizer")
        st.markdown("""
        Optimize sample allocation across multiple dimensions (Region, Size, Industry)
        with constraints on panel and fresh samples.
        """)
        
        # Sidebar configuration
        config = self.render_sidebar()
        
        # Main content
        data = self.render_file_upload()
        
        if data is not None:
            st.divider()
            
            # Scenario configuration
            st.header("🎯 Scenario Configuration")
            
            col1, col2 = st.columns(2)
            
            with col1:
                scenario1_config = self.render_scenario_config(1)
            
            with col2:
                scenario2_config = self.render_scenario_config(2)
            
            st.divider()
            
            # Run optimization
            if st.button("🚀 Run Optimization", type="primary", use_container_width=True):
                results = self.run_optimization(
                    data,
                    config,
                    [scenario1_config, scenario2_config]
                )
                
                if results:
                    st.divider()
                    self.render_results(results)
                    st.divider()
                    self.render_export_options(results)
            
            # Show existing results if available
            elif st.session_state.optimization_results:
                st.divider()
                self.render_results(st.session_state.optimization_results)
                st.divider()
                self.render_export_options(st.session_state.optimization_results)


def main():
    """Application entry point."""
    try:
        app = SampleAllocationApp()
        app.run()
    except Exception as e:
        logger.error(f"Application error: {e}", exc_info=True)
        st.error(f"❌ Application error: {str(e)}")
        st.info("Please refresh the page and try again.")


if __name__ == "__main__":
    main()
