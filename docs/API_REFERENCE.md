# API Reference

Detailed API documentation for the Sample Allocation Optimizer.

## Table of Contents

- [Core Modules](#core-modules)
- [Optimization](#optimization)
- [Utilities](#utilities)
- [Visualization](#visualization)
- [Configuration](#configuration)

---

## Core Modules

### app.py

Main Streamlit application entry point.

#### `SampleAllocationApp`

Main application class.

**Methods:**

##### `__init__()`
Initialize the application with necessary components.

```python
app = SampleAllocationApp()
```

##### `run()`
Main application entry point. Renders the entire UI and handles user interactions.

```python
app.run()
```

##### `render_sidebar() -> Dict`
Render sidebar configuration and return config dict.

**Returns:**
- `Dict`: Configuration dictionary with user-selected parameters

##### `render_file_upload() -> Optional[pd.DataFrame]`
Render file upload section and return loaded data.

**Returns:**
- `pd.DataFrame` or `None`: Loaded data if successful

##### `run_optimization(data, config, scenario_configs) -> Optional[Dict]`
Run optimization for all scenarios.

**Parameters:**
- `data` (pd.DataFrame): Input data
- `config` (Dict): Global configuration
- `scenario_configs` (List[Dict]): Scenario-specific configurations

**Returns:**
- `Dict` or `None`: Optimization results

---

## Optimization

### optimizer.py

Core optimization functionality using CVXPY.

#### `SampleOptimizer`

Optimizer for sample allocation across multiple dimensions.

**Constructor:**

```python
optimizer = SampleOptimizer(config=None)
```

**Parameters:**
- `config` (AppConfig, optional): Application configuration

**Methods:**

##### `optimize(data, config) -> Dict`

Main optimization method.

**Parameters:**
- `data` (pd.DataFrame): Input data with columns:
  - Region, Size, Industry, Population, PanelAvailable
- `config` (Dict): Configuration dictionary with keys:
  - `target_sample` (int): Target total sample size
  - `min_fresh_pct` (float): Minimum fresh sample percentage (0-1)
  - `max_panel_pct` (float): Maximum panel sample percentage (0-1)
  - `tolerance` (float): Tolerance for target deviation (0-1)
  - `confidence_level` (float): Confidence level for calculations
  - `margin_of_error` (float): Desired margin of error
  - `solver` (str): Solver name ('ECOS', 'SCS', 'GLPK_MI')
  - `enforce_region_mins` (bool): Enforce region minimums
  - `enforce_size_mins` (bool): Enforce size minimums
  - `enforce_industry_mins` (bool): Enforce industry minimums

**Returns:**
- `Dict`: Results dictionary containing:
  - `success` (bool): Whether optimization succeeded
  - `status` (str): Solver status
  - `total_sample` (int): Total allocated sample
  - `panel_sample` (int): Panel sample allocated
  - `fresh_sample` (int): Fresh sample allocated
  - `panel_pct` (float): Panel percentage
  - `fresh_pct` (float): Fresh percentage
  - `df_long` (pd.DataFrame): Detailed results by cell
  - `pivot_panel` (pd.DataFrame): Panel allocation pivot
  - `pivot_fresh` (pd.DataFrame): Fresh allocation pivot
  - `df_combined` (pd.DataFrame): Combined samples and base weights
  - `region_totals` (pd.DataFrame): Totals by region
  - `size_totals` (pd.DataFrame): Totals by size
  - `industry_totals` (pd.DataFrame): Totals by industry
  - `solve_time` (float): Time taken to solve
  - `warnings` (List[str], optional): Any warnings

**Example:**

```python
from optimization.optimizer import SampleOptimizer
import pandas as pd

# Prepare data
data = pd.DataFrame({
    'Region': ['North', 'South', 'East', 'West'],
    'Size': ['Small', 'Medium', 'Large', 'Small'],
    'Industry': ['Tech', 'Finance', 'Healthcare', 'Tech'],
    'Population': [1000, 1500, 2000, 800],
    'PanelAvailable': [100, 150, 200, 80]
})

# Configure
config = {
    'target_sample': 500,
    'min_fresh_pct': 0.20,
    'max_panel_pct': 0.80,
    'tolerance': 0.05,
    'confidence_level': 0.95,
    'margin_of_error': 0.05,
    'solver': 'ECOS'
}

# Optimize
optimizer = SampleOptimizer()
result = optimizer.optimize(data, config)

if result['success']:
    print(f"Total: {result['total_sample']}")
    print(result['df_long'])
```

##### `compute_n_infinity(z_score, margin_of_error, p=0.5) -> float`

Compute infinite population sample size.

**Parameters:**
- `z_score` (float): Z-score for confidence level
- `margin_of_error` (float): Desired margin of error
- `p` (float, optional): Expected proportion (default 0.5)

**Returns:**
- `float`: Required sample size for infinite population

**Formula:**
```
n = (z² × p × (1-p)) / e²
```

**Example:**

```python
n_inf = optimizer.compute_n_infinity(
    z_score=1.96,
    margin_of_error=0.05
)
# n_inf ≈ 385
```

##### `compute_fpc_min(N, n_infinity) -> float`

Compute finite population corrected minimum sample size.

**Parameters:**
- `N` (int): Population size
- `n_infinity` (float): Infinite population sample size

**Returns:**
- `float`: Minimum sample size adjusted for finite population

**Formula:**
```
n = n_inf / (1 + n_inf/N)
```

**Example:**

```python
n_min = optimizer.compute_fpc_min(N=5000, n_infinity=385)
# n_min ≈ 357
```

##### `calculate_minimums(data, config) -> Dict[str, pd.DataFrame]`

Calculate minimum sample sizes for all dimensions.

**Parameters:**
- `data` (pd.DataFrame): Input data
- `config` (Dict): Configuration with confidence_level and margin_of_error

**Returns:**
- `Dict[str, pd.DataFrame]`: Dictionary with keys:
  - `'region'`: Region minimums DataFrame
  - `'size'`: Size minimums DataFrame
  - `'industry'`: Industry minimums DataFrame

Each DataFrame contains columns: dimension name, Population, MinNeeded

---

## Utilities

### file_handlers.py

File handling and validation utilities.

#### `FileHandler`

Handler for file operations.

**Constructor:**

```python
handler = FileHandler(config=None)
```

**Methods:**

##### `load_excel(file, sheet_name=None) -> pd.DataFrame`

Load and validate Excel file.

**Parameters:**
- `file`: File object or path
- `sheet_name` (str, optional): Sheet name to load

**Returns:**
- `pd.DataFrame`: Validated and cleaned data

**Raises:**
- `FileValidationError`: If validation fails

**Example:**

```python
from utils.file_handlers import FileHandler

handler = FileHandler()
data = handler.load_excel('input.xlsx')
```

##### `validate_dataframe(df) -> None`

Validate DataFrame structure and content.

**Parameters:**
- `df` (pd.DataFrame): DataFrame to validate

**Raises:**
- `FileValidationError`: If validation fails

**Validation Checks:**
- Required columns present
- No empty data
- Numeric columns are numeric
- No negative values in numeric columns
- No missing values in required columns

##### `clean_dataframe(df) -> pd.DataFrame`

Clean and standardize DataFrame.

**Parameters:**
- `df` (pd.DataFrame): DataFrame to clean

**Returns:**
- `pd.DataFrame`: Cleaned DataFrame

**Cleaning Operations:**
- Strip whitespace from strings
- Ensure proper numeric types
- Remove completely empty rows
- Sort by dimensions

##### `get_data_summary(df) -> Dict`

Get summary statistics for data.

**Parameters:**
- `df` (pd.DataFrame): Data to summarize

**Returns:**
- `Dict`: Summary statistics including:
  - `total_rows`: Number of rows
  - `total_population`: Sum of population
  - `total_panel_available`: Sum of panel available
  - `unique_regions`: Number of unique regions
  - `unique_sizes`: Number of unique sizes
  - `unique_industries`: Number of unique industries

### session_manager.py

Session persistence utilities.

#### `SessionManager`

Manager for saving and loading sessions.

**Constructor:**

```python
manager = SessionManager(config=None)
```

**Methods:**

##### `generate_session_id() -> str`

Generate unique session ID.

**Returns:**
- `str`: UUID-based session ID

##### `save_session(session_id, data, metadata=None) -> bool`

Save session to JSON file.

**Parameters:**
- `session_id` (str): Unique identifier
- `data` (Dict): Data to save
- `metadata` (Dict, optional): Metadata

**Returns:**
- `bool`: True if successful

##### `load_session(session_id) -> Optional[Dict]`

Load session from JSON file.

**Parameters:**
- `session_id` (str): Session identifier

**Returns:**
- `Dict` or `None`: Session data if found

##### `list_sessions() -> List[Dict]`

List all saved sessions.

**Returns:**
- `List[Dict]`: List of session metadata

---

## Visualization

### excel_generator.py

Excel report generation with formatting.

#### `ExcelGenerator`

Generator for formatted Excel reports.

**Constructor:**

```python
generator = ExcelGenerator(config=None)
```

**Methods:**

##### `generate(results, input_data) -> BytesIO`

Generate comprehensive Excel report.

**Parameters:**
- `results` (Dict): Optimization results for all scenarios
- `input_data` (pd.DataFrame): Original input data

**Returns:**
- `BytesIO`: Excel file as bytes

**Example:**

```python
from visualization.excel_generator import ExcelGenerator

generator = ExcelGenerator()
excel_file = generator.generate(results, input_data)

# Save to file
with open('report.xlsx', 'wb') as f:
    f.write(excel_file.getvalue())
```

### html_generator.py

HTML report generation.

#### `HTMLGenerator`

Generator for HTML reports.

**Constructor:**

```python
generator = HTMLGenerator()
```

**Methods:**

##### `generate(results, input_data) -> bytes`

Generate HTML report.

**Parameters:**
- `results` (Dict): Optimization results
- `input_data` (pd.DataFrame): Original input data

**Returns:**
- `bytes`: HTML file as bytes

---

## Configuration

### settings.py

Configuration classes and constants.

#### `AppConfig`

Application configuration settings.

**Attributes:**

- `DEFAULT_CONFIDENCE_LEVEL` (float): Default confidence level (0.95)
- `DEFAULT_MARGIN_OF_ERROR` (float): Default margin of error (0.05)
- `DEFAULT_P` (float): Default proportion for max variability (0.5)
- `Z_SCORES` (Dict[float, float]): Z-score mapping
- `DEFAULT_SOLVER` (str): Default solver ('ECOS')
- `DEFAULT_MAX_ITERATIONS` (int): Default max iterations (10000)
- `REQUIRED_COLUMNS` (List[str]): Required data columns
- `DIMENSION_COLUMNS` (List[str]): Dimension columns
- `NUMERIC_COLUMNS` (List[str]): Numeric columns

**Methods:**

##### `get_z_score(confidence_level) -> float`

Get z-score for confidence level.

**Parameters:**
- `confidence_level` (float): Confidence level (0.90-0.999)

**Returns:**
- `float`: Corresponding z-score

**Raises:**
- `ValueError`: If confidence level out of range

##### `validate_config(config) -> List[str]`

Validate configuration parameters.

**Parameters:**
- `config` (Dict): Configuration to validate

**Returns:**
- `List[str]`: List of error messages (empty if valid)

#### `SolverConfig`

Solver-specific configuration.

**Attributes:**

- `ECOS_OPTIONS` (Dict): ECOS solver options
- `SCS_OPTIONS` (Dict): SCS solver options
- `GLPK_OPTIONS` (Dict): GLPK solver options

**Methods:**

##### `get_solver_options(solver_name) -> Dict`

Get options for specific solver.

**Parameters:**
- `solver_name` (str): Solver name

**Returns:**
- `Dict`: Solver options

---

## Type Definitions

### Common Types

```python
from typing import Dict, List, Optional, Any, Tuple
import pandas as pd

# Configuration dictionary
ConfigDict = Dict[str, Any]

# Results dictionary
ResultDict = Dict[str, Any]

# Minimums dictionary
MinimumsDict = Dict[str, pd.DataFrame]

# Session data
SessionData = Dict[str, Any]
```

---

## Error Handling

### Custom Exceptions

#### `OptimizationError`

Raised when optimization fails.

```python
from optimization.optimizer import OptimizationError

try:
    result = optimizer.optimize(data, config)
except OptimizationError as e:
    print(f"Optimization failed: {e}")
```

#### `FileValidationError`

Raised when file validation fails.

```python
from utils.file_handlers import FileValidationError

try:
    data = handler.load_excel(file)
except FileValidationError as e:
    print(f"Invalid file: {e}")
```

---

## Best Practices

### 1. Data Preparation

```python
# Always validate data before optimization
handler = FileHandler()
handler.validate_dataframe(data)
data = handler.clean_dataframe(data)
```

### 2. Configuration Validation

```python
# Validate config before use
config_obj = AppConfig()
errors = config_obj.validate_config(config)
if errors:
    raise ValueError(f"Config errors: {errors}")
```

### 3. Error Handling

```python
# Always check optimization success
result = optimizer.optimize(data, config)
if not result['success']:
    print(f"Optimization failed: {result.get('message')}")
    # Handle failure
else:
    # Process results
    df_results = result['df_long']
```

### 4. Resource Management

```python
# Use context managers for files
from visualization.excel_generator import ExcelGenerator

generator = ExcelGenerator()
excel_file = generator.generate(results, data)

with open('report.xlsx', 'wb') as f:
    f.write(excel_file.getvalue())
```

---

## Versioning

**Current Version:** 2.0.0

### Version History

- **2.0.0** (January 2026): Complete refactor with improved architecture
- **1.0.0** (Initial): Original implementation
