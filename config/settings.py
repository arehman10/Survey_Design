"""
Configuration settings for the Sample Allocation Optimizer.

This module contains all configuration constants and default values.
"""

from dataclasses import dataclass
from typing import Dict, List
import numpy as np


@dataclass
class AppConfig:
    """Application configuration settings."""
    
    # Statistical defaults
    DEFAULT_CONFIDENCE_LEVEL: float = 0.95
    DEFAULT_MARGIN_OF_ERROR: float = 0.05
    DEFAULT_P: float = 0.5  # Maximum variability assumption
    
    # Z-score mapping for common confidence levels
    Z_SCORES: Dict[float, float] = None
    
    # Optimization defaults
    DEFAULT_SOLVER: str = 'ECOS'
    DEFAULT_MAX_ITERATIONS: int = 10000
    DEFAULT_TOLERANCE: float = 1e-6
    
    # Constraints defaults
    DEFAULT_MIN_FRESH_PCT: float = 0.20
    DEFAULT_MAX_PANEL_PCT: float = 0.80
    DEFAULT_TOLERANCE_PCT: float = 0.05
    
    # File settings
    SESSIONS_DIR: str = "sessions"
    MAX_UPLOAD_SIZE_MB: int = 50
    
    # Excel formatting
    EXCEL_COLOR_SCALE_LOW: str = "00FF00"  # Green
    EXCEL_COLOR_SCALE_MID: str = "FFFF00"  # Yellow
    EXCEL_COLOR_SCALE_HIGH: str = "FF0000"  # Red
    
    # Required columns in input data
    REQUIRED_COLUMNS: List[str] = None
    
    # Dimension columns
    DIMENSION_COLUMNS: List[str] = None
    
    # Numeric columns
    NUMERIC_COLUMNS: List[str] = None
    
    def __post_init__(self):
        """Initialize default values that require computation."""
        if self.Z_SCORES is None:
            self.Z_SCORES = {
                0.90: 1.645,
                0.95: 1.96,
                0.99: 2.576,
                0.999: 3.291
            }
        
        if self.REQUIRED_COLUMNS is None:
            self.REQUIRED_COLUMNS = [
                'Region',
                'Size',
                'Industry',
                'PanelAvailable',
                'Population'
            ]
        
        if self.DIMENSION_COLUMNS is None:
            self.DIMENSION_COLUMNS = [
                'Region',
                'Size',
                'Industry'
            ]
        
        if self.NUMERIC_COLUMNS is None:
            self.NUMERIC_COLUMNS = [
                'PanelAvailable',
                'Population'
            ]
    
    def get_z_score(self, confidence_level: float) -> float:
        """
        Get z-score for a given confidence level.
        
        Args:
            confidence_level: Confidence level (e.g., 0.95 for 95%)
            
        Returns:
            Z-score value
            
        Raises:
            ValueError: If confidence level is not supported
        """
        if confidence_level in self.Z_SCORES:
            return self.Z_SCORES[confidence_level]
        
        # Interpolate for non-standard values
        levels = sorted(self.Z_SCORES.keys())
        if confidence_level < min(levels) or confidence_level > max(levels):
            raise ValueError(
                f"Confidence level {confidence_level} out of range "
                f"[{min(levels)}, {max(levels)}]"
            )
        
        # Linear interpolation
        for i in range(len(levels) - 1):
            if levels[i] <= confidence_level <= levels[i + 1]:
                x0, x1 = levels[i], levels[i + 1]
                y0, y1 = self.Z_SCORES[x0], self.Z_SCORES[x1]
                return y0 + (y1 - y0) * (confidence_level - x0) / (x1 - x0)
        
        raise ValueError(f"Could not interpolate z-score for {confidence_level}")
    
    def validate_config(self, config: Dict) -> List[str]:
        """
        Validate configuration parameters.
        
        Args:
            config: Configuration dictionary
            
        Returns:
            List of validation errors (empty if valid)
        """
        errors = []
        
        # Validate percentages
        if 'min_fresh_pct' in config:
            if not 0 <= config['min_fresh_pct'] <= 1:
                errors.append("min_fresh_pct must be between 0 and 1")
        
        if 'max_panel_pct' in config:
            if not 0 <= config['max_panel_pct'] <= 1:
                errors.append("max_panel_pct must be between 0 and 1")
        
        if 'tolerance' in config:
            if not 0 <= config['tolerance'] <= 1:
                errors.append("tolerance must be between 0 and 1")
        
        # Validate sample size
        if 'target_sample' in config:
            if config['target_sample'] <= 0:
                errors.append("target_sample must be positive")
        
        # Validate consistency
        if 'min_fresh_pct' in config and 'max_panel_pct' in config:
            if config['min_fresh_pct'] + config['max_panel_pct'] > 1:
                errors.append(
                    "min_fresh_pct + max_panel_pct cannot exceed 1"
                )
        
        return errors


@dataclass
class SolverConfig:
    """Solver-specific configuration."""
    
    # Solver options
    ECOS_OPTIONS: Dict = None
    SCS_OPTIONS: Dict = None
    GLPK_OPTIONS: Dict = None
    
    def __post_init__(self):
        """Initialize solver options."""
        if self.ECOS_OPTIONS is None:
            self.ECOS_OPTIONS = {
                'max_iters': 10000,
                'abstol': 1e-7,
                'reltol': 1e-6,
                'feastol': 1e-7
            }
        
        if self.SCS_OPTIONS is None:
            self.SCS_OPTIONS = {
                'max_iters': 10000,
                'eps': 1e-5,
                'alpha': 1.5
            }
        
        if self.GLPK_OPTIONS is None:
            self.GLPK_OPTIONS = {
                'msg_lev': 'GLP_MSG_OFF',
                'tm_lim': 60000  # 60 seconds
            }
    
    def get_solver_options(self, solver_name: str) -> Dict:
        """
        Get options for a specific solver.
        
        Args:
            solver_name: Name of the solver
            
        Returns:
            Dictionary of solver options
        """
        solver_map = {
            'ECOS': self.ECOS_OPTIONS,
            'SCS': self.SCS_OPTIONS,
            'GLPK_MI': self.GLPK_OPTIONS
        }
        
        return solver_map.get(solver_name, {})


# Create singleton instances
app_config = AppConfig()
solver_config = SolverConfig()
