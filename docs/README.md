# Sample Allocation Optimizer

A comprehensive tool for optimizing sample allocation across multiple dimensions (Region, Size, Industry) while balancing panel and fresh samples with statistical constraints.

## 🌟 Features

- **Multi-Dimensional Optimization**: Allocate samples across Region × Size × Industry combinations
- **Dual Scenario Analysis**: Compare two different optimization scenarios side-by-side
- **Statistical Rigor**: Implements finite population correction and confidence intervals
- **Flexible Constraints**: Configure minimum/maximum samples, panel/fresh ratios
- **Professional Reports**: Generate formatted Excel and HTML reports
- **Interactive UI**: Built with Streamlit for ease of use
- **Session Management**: Save and reload optimization sessions
- **Comprehensive Validation**: Robust error handling and data validation

## 📋 Table of Contents

- [Installation](#installation)
- [Quick Start](#quick-start)
- [User Guide](#user-guide)
- [Input Data Format](#input-data-format)
- [Configuration Options](#configuration-options)
- [Understanding Results](#understanding-results)
- [API Documentation](#api-documentation)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)

## 🚀 Installation

### Prerequisites

- Python 3.8 or higher
- pip package manager

### Step 1: Clone or Download

```bash
git clone <repository-url>
cd improved_solver
```

### Step 2: Install Dependencies

```bash
pip install -r requirements.txt
```

**Required packages:**
- streamlit >= 1.28.0
- pandas >= 2.0.0
- numpy >= 1.24.0
- cvxpy >= 1.4.0
- openpyxl >= 3.1.0
- pytest >= 7.4.0 (for testing)

### Step 3: Verify Installation

```bash
python -m pytest tests/ -v
```

## 🎯 Quick Start

### 1. Prepare Your Data

Create an Excel file with the following columns:
- `Region`: Geographic region (e.g., "North", "South")
- `Size`: Business size (e.g., "Small", "Medium", "Large")
- `Industry`: Industry sector (e.g., "Tech", "Finance")
- `Population`: Total population for this segment
- `PanelAvailable`: Number of panel members available

**Example:**
```
Region  | Size   | Industry   | Population | PanelAvailable
North   | Small  | Tech       | 1500       | 200
North   | Small  | Finance    | 1200       | 150
...
```

### 2. Launch the Application

```bash
cd src
streamlit run app.py
```

### 3. Use the Application

1. **Upload Data**: Click "Browse files" and select your Excel file
2. **Configure Scenarios**: Set parameters for Scenario 1 and Scenario 2
3. **Run Optimization**: Click "Run Optimization"
4. **Review Results**: Examine results in the tabs
5. **Export Reports**: Download Excel or HTML reports

## 📖 User Guide

### Understanding the Interface

#### Sidebar Configuration

**Statistical Parameters:**
- **Confidence Level**: The confidence level for sample size calculations (90-99%)
  - Higher values = larger required samples
  - Standard: 95% (z-score = 1.96)
  
- **Margin of Error**: Acceptable error margin (1-10%)
  - Lower values = larger required samples
  - Standard: 5%

**Optimization Settings:**
- **Solver**: Algorithm to use (ECOS recommended for most cases)
  - ECOS: Fast, good for medium problems
  - SCS: Scalable, good for large problems
  - GLPK_MI: Integer programming solver
  
- **Max Iterations**: Maximum solver iterations
  - Increase if optimization fails to converge

#### Scenario Configuration

For each scenario, configure:

**Basic Settings:**
- **Target Sample Size**: Desired total sample (e.g., 1000)
- **Min Fresh %**: Minimum percentage of fresh samples (0-100%)
- **Max Panel %**: Maximum percentage of panel samples (0-100%)
- **Tolerance %**: Acceptable deviation from target (0-20%)

**Advanced Constraints:**
- **Enforce industry minimums**: Ensure minimum samples per industry
- **Enforce region minimums**: Ensure minimum samples per region
- **Enforce size minimums**: Ensure minimum samples per size

### Optimization Process

The optimizer:

1. **Calculates Minimums**: Uses finite population correction to determine minimum required samples for each dimension
2. **Builds Problem**: Creates a constrained optimization problem with your requirements
3. **Solves**: Uses convex optimization to find the best allocation
4. **Validates**: Checks that all constraints are satisfied
5. **Generates Reports**: Creates detailed output tables

### Reading Results

#### Summary Tab

- **Total Sample**: Total allocated samples
- **Panel Sample**: Samples from existing panel
- **Fresh Sample**: New samples needed
- **Panel %** / **Fresh %**: Percentage breakdown
- **Optimization Time**: How long the solver took

#### Detailed Results Tab

Shows the complete allocation table with:
- All dimension combinations (Region × Size × Industry)
- Panel and Fresh samples for each cell
- Total samples per cell
- Base weights (Population / Sample)

#### Comparison Tab

Compares Scenario 1 vs Scenario 2:
- Differences in sample allocation
- Panel/Fresh distribution changes
- Base weight variations

#### Diagnostics Tab

Technical information:
- Solver status (optimal, optimal_inaccurate, infeasible, etc.)
- Number of iterations
- Objective value
- Warnings (if any)

## 📊 Input Data Format

### Required Columns

| Column | Type | Description | Example |
|--------|------|-------------|---------|
| Region | String | Geographic region | "Northeast", "West" |
| Size | String | Business size | "Small", "Medium", "Large" |
| Industry | String | Industry sector | "Technology", "Healthcare" |
| Population | Integer | Total population in segment | 5000 |
| PanelAvailable | Integer | Available panel members | 250 |

### Data Requirements

- **No missing values** in required columns
- **Positive numbers** for Population and PanelAvailable
- **Consistent naming** within dimensions
- **Complete coverage**: All combinations should be present

### Example Data

Download sample data: [sample_data.xlsx](examples/sample_data.xlsx)

## ⚙️ Configuration Options

### Statistical Parameters

```python
{
    'confidence_level': 0.95,      # 90%, 95%, 99%, or 99.9%
    'margin_of_error': 0.05,       # 1% - 10%
}
```

**How it works:**
- Higher confidence = larger samples needed
- Lower margin of error = larger samples needed
- Formula: n = (z² × p × (1-p)) / e²
- With FPC: n_final = n_inf / (1 + n_inf/N)

### Optimization Constraints

```python
{
    'target_sample': 1000,          # Target total sample
    'min_fresh_pct': 0.20,          # At least 20% fresh
    'max_panel_pct': 0.80,          # At most 80% panel
    'tolerance': 0.05,              # ±5% of target OK
    'enforce_region_mins': True,    # Enforce region minimums
    'enforce_size_mins': True,      # Enforce size minimums
    'enforce_industry_mins': True,  # Enforce industry minimums
}
```

### Solver Options

```python
{
    'solver': 'ECOS',               # ECOS, SCS, or GLPK_MI
    'max_iterations': 10000,        # Maximum iterations
}
```

**Solver Selection Guide:**
- **ECOS**: Best for most cases, fast and accurate
- **SCS**: Better for very large problems (1000+ variables)
- **GLPK_MI**: For strict integer constraints

## 📈 Understanding Results

### Base Weights

**Base Weight = Population / Sample**

- Indicates how many population units each sample represents
- Lower is better (more representative)
- Should be similar across cells (homogeneous weighting)
- Color-coded in reports: Green (low) → Yellow (medium) → Red (high)

### Optimization Status

- **optimal**: Perfect solution found ✓
- **optimal_inaccurate**: Good solution found (small numerical issues) ✓
- **infeasible**: No solution satisfies all constraints ✗
- **unbounded**: Problem not properly constrained ✗
- **solver_error**: Technical problem ✗

### Common Issues

**"Infeasible" Status:**
- Constraints are too restrictive
- Try: Increase tolerance, relax minimums, increase target sample

**"High Base Weights":**
- Some segments undersampled
- Try: Increase target sample or adjust constraints

**"Long Solve Time":**
- Problem is complex or large
- Try: Use SCS solver, reduce constraints, increase max iterations

## 🔧 API Documentation

### Using as a Library

```python
from optimization.optimizer import SampleOptimizer
from utils.file_handlers import FileHandler
import pandas as pd

# Load data
handler = FileHandler()
data = handler.load_excel('input_data.xlsx')

# Configure optimization
config = {
    'target_sample': 1000,
    'min_fresh_pct': 0.20,
    'max_panel_pct': 0.80,
    'tolerance': 0.05,
    'confidence_level': 0.95,
    'margin_of_error': 0.05,
    'solver': 'ECOS'
}

# Run optimization
optimizer = SampleOptimizer()
result = optimizer.optimize(data, config)

# Check results
if result['success']:
    print(f"Total Sample: {result['total_sample']}")
    print(f"Panel: {result['panel_sample']}")
    print(f"Fresh: {result['fresh_sample']}")
    
    # Access detailed results
    df_results = result['df_long']
    print(df_results.head())
else:
    print(f"Optimization failed: {result['message']}")
```

### Key Classes

#### SampleOptimizer

```python
optimizer = SampleOptimizer(config=None)
result = optimizer.optimize(data, config)
```

**Methods:**
- `optimize(data, config)`: Run optimization
- `calculate_minimums(data, config)`: Calculate minimum samples
- `compute_n_infinity(z_score, moe, p)`: Infinite population sample size
- `compute_fpc_min(N, n_inf)`: Finite population correction

#### FileHandler

```python
handler = FileHandler(config=None)
data = handler.load_excel(file, sheet_name=None)
handler.validate_dataframe(data)
```

**Methods:**
- `load_excel(file)`: Load and validate Excel file
- `validate_dataframe(df)`: Validate data structure
- `clean_dataframe(df)`: Clean and standardize data
- `get_data_summary(df)`: Get summary statistics

#### ExcelGenerator

```python
from visualization.excel_generator import ExcelGenerator

generator = ExcelGenerator()
excel_file = generator.generate(results, input_data)
```

## 🐛 Troubleshooting

### Common Errors

**"Missing required columns"**
- Ensure your Excel has all required columns
- Check spelling: "Region" not "region"
- No extra spaces in column names

**"File contains no data rows"**
- Excel sheet is empty
- Wrong sheet selected
- Header row missing

**"Column contains non-numeric values"**
- Population and PanelAvailable must be numbers
- Remove any text, commas, or currency symbols

**"Solver status: infeasible"**
- Constraints are contradictory
- Not enough panel available
- Target too high or too low
- Solution: Relax constraints or adjust target

**"Memory Error"**
- Dataset too large
- Solution: Reduce data size or increase system memory

### Getting Help

1. Check logs in console
2. Review configuration for errors
3. Try with sample data
4. Contact support with:
   - Error message
   - Input data (if possible)
   - Configuration used

## 🧪 Testing

Run the test suite:

```bash
# All tests
pytest tests/ -v

# Specific test file
pytest tests/test_optimizer.py -v

# With coverage
pytest tests/ --cov=src --cov-report=html
```

## 📝 Contributing

### Development Setup

```bash
# Install dev dependencies
pip install -r requirements-dev.txt

# Run linting
flake8 src/
black src/

# Run type checking
mypy src/
```

### Code Style

- Follow PEP 8
- Use type hints
- Add docstrings to all functions
- Write tests for new features

## 📄 License

This project is proprietary software. All rights reserved.

## 🙏 Acknowledgments

- Built with [Streamlit](https://streamlit.io/)
- Optimization powered by [CVXPY](https://www.cvxpy.org/)
- Data handling with [Pandas](https://pandas.pydata.org/)

## 📞 Support

For questions or issues:
- Email: support@example.com
- Documentation: https://docs.example.com
- Issue Tracker: https://github.com/example/issues

---

**Version**: 2.0.0  
**Last Updated**: January 2026
