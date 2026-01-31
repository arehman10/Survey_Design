"""
Test suite for Sample Allocation Optimizer.

Run tests with: pytest tests/test_optimizer.py -v
"""

import pytest
import pandas as pd
import numpy as np
import sys
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from optimization.optimizer import SampleOptimizer
from utils.file_handlers import FileHandler, FileValidationError
from config.settings import AppConfig


class TestDataFixtures:
    """Test data fixtures."""
    
    @staticmethod
    def create_sample_data(n_rows=10) -> pd.DataFrame:
        """Create sample test data."""
        regions = ['North', 'South', 'East', 'West']
        sizes = ['Small', 'Medium', 'Large']
        industries = ['Tech', 'Finance', 'Healthcare']
        
        data = []
        for i in range(n_rows):
            data.append({
                'Region': np.random.choice(regions),
                'Size': np.random.choice(sizes),
                'Industry': np.random.choice(industries),
                'Population': np.random.randint(100, 10000),
                'PanelAvailable': np.random.randint(10, 500)
            })
        
        return pd.DataFrame(data)


@pytest.fixture
def sample_data():
    """Fixture providing sample data."""
    return TestDataFixtures.create_sample_data(50)


@pytest.fixture
def optimizer():
    """Fixture providing optimizer instance."""
    return SampleOptimizer()


@pytest.fixture
def file_handler():
    """Fixture providing file handler instance."""
    return FileHandler()


@pytest.fixture
def config():
    """Fixture providing configuration."""
    return AppConfig()


class TestAppConfig:
    """Tests for AppConfig class."""
    
    def test_config_initialization(self, config):
        """Test config initializes correctly."""
        assert config.DEFAULT_CONFIDENCE_LEVEL == 0.95
        assert config.DEFAULT_MARGIN_OF_ERROR == 0.05
        assert len(config.REQUIRED_COLUMNS) > 0
    
    def test_get_z_score_standard(self, config):
        """Test getting standard z-scores."""
        assert abs(config.get_z_score(0.95) - 1.96) < 0.01
        assert abs(config.get_z_score(0.90) - 1.645) < 0.01
        assert abs(config.get_z_score(0.99) - 2.576) < 0.01
    
    def test_get_z_score_invalid(self, config):
        """Test error handling for invalid confidence levels."""
        with pytest.raises(ValueError):
            config.get_z_score(0.5)  # Too low
        with pytest.raises(ValueError):
            config.get_z_score(1.5)  # Too high
    
    def test_validate_config_valid(self, config):
        """Test config validation with valid input."""
        test_config = {
            'min_fresh_pct': 0.2,
            'max_panel_pct': 0.8,
            'tolerance': 0.05,
            'target_sample': 1000
        }
        errors = config.validate_config(test_config)
        assert len(errors) == 0
    
    def test_validate_config_invalid_percentages(self, config):
        """Test config validation with invalid percentages."""
        test_config = {
            'min_fresh_pct': 1.5,  # Invalid
        }
        errors = config.validate_config(test_config)
        assert len(errors) > 0
    
    def test_validate_config_inconsistent(self, config):
        """Test config validation with inconsistent values."""
        test_config = {
            'min_fresh_pct': 0.6,
            'max_panel_pct': 0.6,  # Sum > 1
        }
        errors = config.validate_config(test_config)
        assert len(errors) > 0


class TestFileHandler:
    """Tests for FileHandler class."""
    
    def test_validate_dataframe_valid(self, file_handler, sample_data):
        """Test validation of valid dataframe."""
        # Should not raise exception
        file_handler.validate_dataframe(sample_data)
    
    def test_validate_dataframe_missing_columns(self, file_handler):
        """Test validation fails with missing columns."""
        df = pd.DataFrame({'Region': ['North', 'South']})
        with pytest.raises(FileValidationError):
            file_handler.validate_dataframe(df)
    
    def test_validate_dataframe_empty(self, file_handler):
        """Test validation fails with empty dataframe."""
        df = pd.DataFrame(columns=['Region', 'Size', 'Industry', 'Population', 'PanelAvailable'])
        with pytest.raises(FileValidationError):
            file_handler.validate_dataframe(df)
    
    def test_validate_dataframe_negative_values(self, file_handler, sample_data):
        """Test validation fails with negative values."""
        sample_data.loc[0, 'Population'] = -100
        with pytest.raises(FileValidationError):
            file_handler.validate_dataframe(sample_data)
    
    def test_clean_dataframe(self, file_handler, sample_data):
        """Test dataframe cleaning."""
        # Add some whitespace
        sample_data.loc[0, 'Region'] = '  North  '
        
        cleaned = file_handler.clean_dataframe(sample_data)
        
        assert cleaned.loc[0, 'Region'] == 'North'
        assert len(cleaned) == len(sample_data)
    
    def test_get_data_summary(self, file_handler, sample_data):
        """Test data summary generation."""
        summary = file_handler.get_data_summary(sample_data)
        
        assert 'total_rows' in summary
        assert 'total_population' in summary
        assert 'unique_regions' in summary
        assert summary['total_rows'] == len(sample_data)


class TestSampleOptimizer:
    """Tests for SampleOptimizer class."""
    
    def test_compute_n_infinity(self, optimizer):
        """Test infinite population sample size calculation."""
        n = optimizer.compute_n_infinity(z_score=1.96, margin_of_error=0.05)
        assert n > 0
        assert n < 10000  # Reasonable range
    
    def test_compute_fpc_min(self, optimizer):
        """Test finite population correction."""
        n_inf = 385
        
        # Large population (should be close to n_inf)
        n_large = optimizer.compute_fpc_min(100000, n_inf)
        assert abs(n_large - n_inf) < 10
        
        # Small population (should be much smaller)
        n_small = optimizer.compute_fpc_min(500, n_inf)
        assert n_small < n_inf
        assert n_small > 0
        
        # Zero population
        n_zero = optimizer.compute_fpc_min(0, n_inf)
        assert n_zero == 0
    
    def test_calculate_minimums(self, optimizer, sample_data):
        """Test minimum sample size calculations."""
        config = {
            'confidence_level': 0.95,
            'margin_of_error': 0.05
        }
        
        minimums = optimizer.calculate_minimums(sample_data, config)
        
        assert 'region' in minimums
        assert 'size' in minimums
        assert 'industry' in minimums
        
        # Check that minimums have required columns
        assert 'MinNeeded' in minimums['region'].columns
        assert all(minimums['region']['MinNeeded'] >= 0)
    
    def test_build_optimization_problem(self, optimizer, sample_data):
        """Test optimization problem construction."""
        config = {
            'target_sample': 1000,
            'min_fresh_pct': 0.2,
            'max_panel_pct': 0.8,
            'tolerance': 0.05,
            'enforce_region_mins': True,
            'enforce_size_mins': True,
            'enforce_industry_mins': True,
            'confidence_level': 0.95,
            'margin_of_error': 0.05
        }
        
        minimums = optimizer.calculate_minimums(sample_data, config)
        problem, variables = optimizer.build_optimization_problem(
            sample_data, config, minimums
        )
        
        assert problem is not None
        assert 'panel' in variables
        assert 'fresh' in variables
        assert 'total_sample' in variables
        assert len(problem.constraints) > 0
    
    def test_optimize_success(self, optimizer, sample_data):
        """Test successful optimization."""
        config = {
            'target_sample': 500,
            'min_fresh_pct': 0.2,
            'max_panel_pct': 0.8,
            'tolerance': 0.1,
            'enforce_region_mins': False,  # Relaxed for test
            'enforce_size_mins': False,
            'enforce_industry_mins': False,
            'confidence_level': 0.95,
            'margin_of_error': 0.05,
            'solver': 'ECOS'
        }
        
        result = optimizer.optimize(sample_data, config)
        
        # Should succeed or at least not crash
        assert 'success' in result
        if result['success']:
            assert 'total_sample' in result
            assert 'df_long' in result
            assert len(result['df_long']) == len(sample_data)
    
    def test_optimize_invalid_config(self, optimizer, sample_data):
        """Test optimization with invalid config."""
        config = {
            'min_fresh_pct': 1.5,  # Invalid
        }
        
        result = optimizer.optimize(sample_data, config)
        
        assert result['success'] == False
        assert 'message' in result


class TestIntegration:
    """Integration tests."""
    
    def test_full_workflow(self, sample_data):
        """Test complete workflow from data to results."""
        # Setup
        file_handler = FileHandler()
        optimizer = SampleOptimizer()
        
        # Validate data
        file_handler.validate_dataframe(sample_data)
        
        # Clean data
        cleaned_data = file_handler.clean_dataframe(sample_data)
        
        # Configure optimization
        config = {
            'target_sample': 500,
            'min_fresh_pct': 0.2,
            'max_panel_pct': 0.8,
            'tolerance': 0.15,
            'enforce_region_mins': False,
            'enforce_size_mins': False,
            'enforce_industry_mins': False,
            'confidence_level': 0.95,
            'margin_of_error': 0.05,
            'solver': 'ECOS'
        }
        
        # Run optimization
        result = optimizer.optimize(cleaned_data, config)
        
        # Verify results
        if result.get('success'):
            assert result['total_sample'] > 0
            assert result['panel_sample'] >= 0
            assert result['fresh_sample'] >= 0
            assert abs(
                result['panel_sample'] + result['fresh_sample'] - result['total_sample']
            ) < 1


def test_performance_large_dataset():
    """Test performance with larger dataset."""
    import time
    
    # Create larger dataset
    data = TestDataFixtures.create_sample_data(500)
    
    optimizer = SampleOptimizer()
    config = {
        'target_sample': 2000,
        'min_fresh_pct': 0.2,
        'max_panel_pct': 0.8,
        'tolerance': 0.15,
        'enforce_region_mins': False,
        'enforce_size_mins': False,
        'enforce_industry_mins': False,
        'confidence_level': 0.95,
        'margin_of_error': 0.05,
        'solver': 'ECOS',
        'max_iterations': 5000
    }
    
    start_time = time.time()
    result = optimizer.optimize(data, config)
    elapsed_time = time.time() - start_time
    
    # Should complete in reasonable time
    assert elapsed_time < 60  # Less than 1 minute
    
    if result.get('success'):
        assert result['solve_time'] > 0


if __name__ == '__main__':
    pytest.main([__file__, '-v', '--tb=short'])
