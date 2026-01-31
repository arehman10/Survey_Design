"""
File handling utilities for the Sample Allocation Optimizer.

This module handles reading Excel files, validating data, and basic
data transformations.
"""

import pandas as pd
import numpy as np
from typing import Optional, List, Dict, Any
import logging
from io import BytesIO

from config.settings import AppConfig

logger = logging.getLogger(__name__)


class FileValidationError(Exception):
    """Custom exception for file validation errors."""
    pass


class FileHandler:
    """Handler for file operations including validation and transformation."""
    
    def __init__(self, config: Optional[AppConfig] = None):
        """
        Initialize the file handler.
        
        Args:
            config: Application configuration
        """
        self.config = config or AppConfig()
    
    def load_excel(
        self,
        file,
        sheet_name: Optional[str] = None
    ) -> pd.DataFrame:
        """
        Load and validate Excel file.
        
        Args:
            file: File object or path
            sheet_name: Specific sheet to load (default: first sheet)
            
        Returns:
            Validated DataFrame
            
        Raises:
            FileValidationError: If file validation fails
        """
        try:
            # Read Excel file
            if isinstance(file, (str, BytesIO)):
                df = pd.read_excel(file, sheet_name=sheet_name or 0)
            else:
                df = pd.read_excel(file.read(), sheet_name=sheet_name or 0)
            
            logger.info(f"Loaded {len(df)} rows from Excel file")
            
            # Validate
            self.validate_dataframe(df)
            
            # Clean and transform
            df = self.clean_dataframe(df)
            
            return df
            
        except FileValidationError:
            raise
        except Exception as e:
            logger.error(f"Error loading Excel file: {e}", exc_info=True)
            raise FileValidationError(f"Error loading Excel file: {str(e)}")
    
    def validate_dataframe(self, df: pd.DataFrame) -> None:
        """
        Validate that DataFrame has required structure.
        
        Args:
            df: DataFrame to validate
            
        Raises:
            FileValidationError: If validation fails
        """
        # Check for required columns
        missing_cols = set(self.config.REQUIRED_COLUMNS) - set(df.columns)
        if missing_cols:
            raise FileValidationError(
                f"Missing required columns: {', '.join(missing_cols)}"
            )
        
        # Check for empty DataFrame
        if len(df) == 0:
            raise FileValidationError("File contains no data rows")
        
        # Check numeric columns
        for col in self.config.NUMERIC_COLUMNS:
            if col in df.columns:
                if not pd.api.types.is_numeric_dtype(df[col]):
                    # Try to convert
                    try:
                        df[col] = pd.to_numeric(df[col], errors='coerce')
                        if df[col].isna().any():
                            raise ValueError(f"Contains non-numeric values")
                    except:
                        raise FileValidationError(
                            f"Column '{col}' must contain numeric values"
                        )
        
        # Check for negative values in numeric columns
        for col in self.config.NUMERIC_COLUMNS:
            if col in df.columns and (df[col] < 0).any():
                raise FileValidationError(
                    f"Column '{col}' contains negative values"
                )
        
        # Check for null values in key columns
        for col in self.config.REQUIRED_COLUMNS:
            if df[col].isna().any():
                raise FileValidationError(
                    f"Column '{col}' contains missing values"
                )
        
        logger.info("DataFrame validation passed")
    
    def clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Clean and standardize DataFrame.
        
        Args:
            df: DataFrame to clean
            
        Returns:
            Cleaned DataFrame
        """
        df = df.copy()
        
        # Strip whitespace from string columns
        for col in self.config.DIMENSION_COLUMNS:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
        
        # Ensure numeric columns are proper type
        for col in self.config.NUMERIC_COLUMNS:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # Remove any completely empty rows
        df = df.dropna(how='all')
        
        # Sort by dimensions for consistency
        df = df.sort_values(
            by=self.config.DIMENSION_COLUMNS,
            ignore_index=True
        )
        
        logger.info("DataFrame cleaned successfully")
        return df
    
    def get_data_summary(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        Get summary statistics for the data.
        
        Args:
            df: DataFrame to summarize
            
        Returns:
            Dictionary with summary statistics
        """
        summary = {
            'total_rows': len(df),
            'total_population': int(df['Population'].sum()) if 'Population' in df.columns else 0,
            'total_panel_available': int(df['PanelAvailable'].sum()) if 'PanelAvailable' in df.columns else 0,
            'unique_regions': df['Region'].nunique() if 'Region' in df.columns else 0,
            'unique_sizes': df['Size'].nunique() if 'Size' in df.columns else 0,
            'unique_industries': df['Industry'].nunique() if 'Industry' in df.columns else 0,
        }
        
        # Add dimension breakdowns
        if 'Region' in df.columns:
            summary['regions'] = df['Region'].unique().tolist()
        if 'Size' in df.columns:
            summary['sizes'] = df['Size'].unique().tolist()
        if 'Industry' in df.columns:
            summary['industries'] = df['Industry'].unique().tolist()
        
        return summary
    
    def export_to_excel(
        self,
        dataframes: Dict[str, pd.DataFrame],
        filename: str
    ) -> BytesIO:
        """
        Export multiple dataframes to Excel file.
        
        Args:
            dataframes: Dictionary mapping sheet names to DataFrames
            filename: Output filename (for metadata only)
            
        Returns:
            BytesIO object containing Excel file
        """
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df in dataframes.items():
                # Truncate sheet name if too long (Excel limit is 31 chars)
                sheet_name = sheet_name[:31]
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        output.seek(0)
        logger.info(f"Exported {len(dataframes)} sheets to Excel")
        return output
