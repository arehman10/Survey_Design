# Troubleshooting Guide

Common issues and solutions for the Sample Allocation Optimizer.

## Table of Contents

1. [Installation Issues](#installation-issues)
2. [Data Loading Errors](#data-loading-errors)
3. [Optimization Failures](#optimization-failures)
4. [Performance Issues](#performance-issues)
5. [Output Issues](#output-issues)
6. [UI/UX Issues](#uiux-issues)

---

## Installation Issues

### Problem: Package installation fails

**Symptoms:**
```
ERROR: Could not find a version that satisfies the requirement cvxpy>=1.4.0
```

**Solutions:**

1. **Update pip:**
```bash
python -m pip install --upgrade pip
```

2. **Install dependencies one by one:**
```bash
pip install numpy pandas
pip install cvxpy
pip install streamlit openpyxl
```

3. **Use specific versions:**
```bash
pip install cvxpy==1.4.1
```

4. **Check Python version:**
```bash
python --version  # Should be 3.8 or higher
```

### Problem: CVXPY solver issues

**Symptoms:**
```
SolverError: The solver ECOS is not installed
```

**Solutions:**

1. **Install specific solvers:**
```bash
pip install cvxopt  # For ECOS
pip install scs     # For SCS
pip install swiglpk # For GLPK
```

2. **Try different solver:**
```python
config['solver'] = 'SCS'  # Instead of ECOS
```

---

## Data Loading Errors

### Problem: "Missing required columns"

**Error Message:**
```
FileValidationError: Missing required columns: Population, PanelAvailable
```

**Solutions:**

1. **Check column names** (case-sensitive):
   - Required: Region, Size, Industry, Population, PanelAvailable
   - ❌ Wrong: region, REGION, Regions
   - ✅ Correct: Region

2. **Check for hidden characters:**
```python
# Print actual column names
print(df.columns.tolist())
```

3. **Rename columns programmatically:**
```python
df = df.rename(columns={
    'Regions': 'Region',
    'Business Size': 'Size'
})
```

### Problem: "Column contains non-numeric values"

**Error Message:**
```
FileValidationError: Column 'Population' must contain numeric values
```

**Solutions:**

1. **Remove commas from numbers:**
   - ❌ Wrong: 1,500
   - ✅ Correct: 1500

2. **Remove currency symbols:**
   - ❌ Wrong: $1500
   - ✅ Correct: 1500

3. **Clean in Excel:**
   - Select column → Format Cells → Number → 0 decimals

4. **Clean programmatically:**
```python
df['Population'] = df['Population'].str.replace(',', '').astype(int)
```

### Problem: "File contains no data rows"

**Solutions:**

1. **Check for empty sheets:**
   - Data should start from row 2 (row 1 = headers)
   
2. **Specify correct sheet:**
```python
data = handler.load_excel(file, sheet_name='Data')
```

3. **Check for filters:**
   - Remove any Excel filters that might hide rows

### Problem: "Column contains missing values"

**Error Message:**
```
FileValidationError: Column 'Region' contains missing values
```

**Solutions:**

1. **Fill missing values:**
```python
df['Region'] = df['Region'].fillna('Unknown')
```

2. **Remove rows with missing values:**
```python
df = df.dropna(subset=['Region', 'Size', 'Industry'])
```

3. **Check Excel formulas:**
   - Formulas returning #N/A or #REF! appear as missing values

---

## Optimization Failures

### Problem: "Solver status: infeasible"

**Meaning:** No solution exists that satisfies all constraints

**Common Causes:**

1. **Target too high:**
```
Target: 10000
Available Panel: 500
Max Panel %: 80% (= 8000 panel needed)
→ Need 2000 fresh, but constraints require more
```

**Solution:** Increase target or relax constraints

2. **Contradictory constraints:**
```
Min Fresh %: 50%
Max Panel %: 60%
→ Only 10% flexibility is impossible
```

**Solution:** Widen the gap:
```python
config['min_fresh_pct'] = 0.20  # 20%
config['max_panel_pct'] = 0.80  # 80%
```

3. **Minimums too high:**
```
Region minimums sum: 1500
Target sample: 1000
→ Can't satisfy both
```

**Solution:** 
- Increase target
- Reduce confidence level (lower minimums)
- Disable dimension minimums temporarily:
```python
config['enforce_region_mins'] = False
config['enforce_size_mins'] = False
```

### Problem: "Solver status: unbounded"

**Meaning:** Problem is not properly constrained

**Solutions:**

1. **Add target constraint:**
```python
config['target_sample'] = 1000  # Must specify
```

2. **Add panel/fresh constraints:**
```python
config['min_fresh_pct'] = 0.20
config['max_panel_pct'] = 0.80
```

### Problem: "Solver status: solver_error"

**Solutions:**

1. **Try different solver:**
```python
config['solver'] = 'SCS'  # Instead of ECOS
```

2. **Increase iterations:**
```python
config['max_iterations'] = 50000
```

3. **Simplify problem:**
   - Reduce data size
   - Relax constraints
   - Increase tolerance

### Problem: "optimal_inaccurate" status

**Meaning:** Solution found but with numerical issues

**Impact:** Usually acceptable, but check results carefully

**Solutions:**

1. **Accept the solution:**
   - Usually fine for practical use
   - Check base weights for reasonableness

2. **If unacceptable:**
```python
config['solver'] = 'SCS'  # More robust
config['max_iterations'] = 20000
```

---

## Performance Issues

### Problem: Optimization takes too long (>5 minutes)

**Causes:**

1. **Large dataset:**
   - 1000+ rows
   - Many dimension combinations

2. **Complex constraints:**
   - All minimums enforced
   - Tight tolerances

**Solutions:**

1. **Use faster solver:**
```python
config['solver'] = 'SCS'  # Faster for large problems
```

2. **Relax constraints:**
```python
config['enforce_region_mins'] = False  # Only if acceptable
config['tolerance'] = 0.10  # Increase tolerance
```

3. **Reduce data:**
```python
# Combine small segments
df = df[df['Population'] >= 100]  # Remove tiny segments
```

4. **Optimize iterations:**
```python
config['max_iterations'] = 5000  # Lower if acceptable
```

### Problem: Memory errors

**Error Message:**
```
MemoryError: Unable to allocate array
```

**Solutions:**

1. **Reduce data size:**
```python
# Sample data
df = df.sample(frac=0.5)

# Or aggregate
df = df.groupby(['Region', 'Industry']).agg({
    'Population': 'sum',
    'PanelAvailable': 'sum'
}).reset_index()
```

2. **Use sparse matrices** (advanced):
```python
# In optimizer.py, modify problem construction
# Use sparse=True in cvxpy.Variable()
```

3. **Increase system memory:**
   - Close other applications
   - Use machine with more RAM

---

## Output Issues

### Problem: Base weights are very high (>100)

**Meaning:** Some segments are severely undersampled

**Example:**
```
Region: East, Size: Small
Population: 5000
Sample: 10
Base Weight: 500  ← Each sample represents 500 people!
```

**Solutions:**

1. **Increase total sample:**
```python
config['target_sample'] = 2000  # Instead of 1000
```

2. **Check minimums:**
```python
# Ensure minimums are being enforced
config['enforce_region_mins'] = True
config['enforce_size_mins'] = True
```

3. **Adjust distribution:**
   - Review panel availability
   - Check population distribution

### Problem: Some cells have 0 samples

**Causes:**

1. **Minimums not enforced:**
```python
config['enforce_industry_mins'] = False  # ← Issue
```

2. **Population too small:**
   - Minimum might round to 0

**Solutions:**

1. **Enforce minimums:**
```python
config['enforce_region_mins'] = True
config['enforce_size_mins'] = True
config['enforce_industry_mins'] = True
```

2. **Set manual minimums:**
```python
# In code, add minimum sample constraints
# For each cell: sample >= 1
```

### Problem: Panel/Fresh ratio incorrect

**Example:**
```
Target: 50% fresh
Actual: 35% fresh
```

**Causes:**

1. **Tolerance too loose:**
```python
config['tolerance'] = 0.20  # ±20% is very loose!
```

2. **Insufficient panel:**
   - Not enough panel available
   - Constraints force more fresh

**Solutions:**

1. **Tighten tolerance:**
```python
config['tolerance'] = 0.05  # ±5% is tighter
```

2. **Check availability:**
```python
# Sum of PanelAvailable should be > target × max_panel_pct
total_panel = df['PanelAvailable'].sum()
required = config['target_sample'] * config['max_panel_pct']
if total_panel < required:
    print(f"Insufficient panel: {total_panel} < {required}")
```

---

## UI/UX Issues

### Problem: Streamlit app won't start

**Error Message:**
```
ModuleNotFoundError: No module named 'streamlit'
```

**Solutions:**

1. **Install Streamlit:**
```bash
pip install streamlit
```

2. **Check installation:**
```bash
streamlit --version
```

3. **Use correct Python:**
```bash
python -m streamlit run src/app.py
```

### Problem: File upload fails

**Solutions:**

1. **Check file size:**
   - Default limit: 200MB
   - Increase in `.streamlit/config.toml`:
```toml
[server]
maxUploadSize = 500
```

2. **Check file format:**
   - Must be .xlsx or .xls
   - Not .csv (convert to Excel first)

3. **Try different browser:**
   - Chrome usually works best
   - Clear browser cache

### Problem: Results don't display

**Solutions:**

1. **Check console for errors:**
   - Open browser dev tools (F12)
   - Look for JavaScript errors

2. **Refresh the page:**
```
Ctrl+F5 (hard refresh)
```

3. **Clear session state:**
```python
# In sidebar, add:
if st.button("Reset"):
    st.session_state.clear()
    st.experimental_rerun()
```

### Problem: Download button doesn't work

**Solutions:**

1. **Check browser settings:**
   - Allow pop-ups from localhost
   - Check downloads folder

2. **Try different format:**
   - If Excel fails, try HTML
   - If HTML fails, try Excel

3. **Check file generation:**
```python
# Add debugging
try:
    excel_file = generator.generate(results, data)
    print(f"File size: {len(excel_file.getvalue())} bytes")
except Exception as e:
    print(f"Generation error: {e}")
```

---

## Debugging Tips

### Enable Logging

```python
import logging

logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
```

### Print Intermediate Results

```python
# After each major step
print(f"Data loaded: {len(data)} rows")
print(f"Minimums calculated: {minimums['region']}")
print(f"Problem built: {len(problem.constraints)} constraints")
print(f"Solver status: {problem.status}")
```

### Validate Each Step

```python
# After data loading
assert len(data) > 0, "No data loaded"

# After cleaning
assert data['Population'].min() >= 0, "Negative population"

# After optimization
assert result['success'], f"Failed: {result.get('message')}"
```

### Test with Minimal Data

```python
# Create simple test case
test_data = pd.DataFrame({
    'Region': ['North', 'South'],
    'Size': ['Small', 'Small'],
    'Industry': ['Tech', 'Tech'],
    'Population': [1000, 1000],
    'PanelAvailable': [100, 100]
})

# Test optimization
result = optimizer.optimize(test_data, simple_config)
```

---

## Getting Additional Help

### Before Asking for Help

1. **Check error message** - Read it carefully
2. **Review this guide** - Most issues are covered
3. **Test with sample data** - Isolate the problem
4. **Check logs** - Enable DEBUG logging

### When Asking for Help

Provide:

1. **Exact error message:**
```
Full stack trace, not just "it doesn't work"
```

2. **Configuration used:**
```python
config = {
    'target_sample': 1000,
    'solver': 'ECOS',
    # ... all settings
}
```

3. **Data characteristics:**
```
Rows: 500
Regions: 4
Industries: 6
Total Population: 50,000
Total Panel: 5,000
```

4. **Steps to reproduce:**
```
1. Load data from file X
2. Set config Y
3. Click optimize
4. Error occurs
```

5. **What you've tried:**
```
- Changed solver to SCS
- Increased tolerance to 0.10
- Disabled minimums
```

---

## Emergency Fallbacks

If all else fails:

### 1. Use Manual Allocation

```python
# Simple proportional allocation
df['Sample'] = (df['Population'] / df['Population'].sum() * target_sample).round()
```

### 2. Use Simpler Constraints

```python
config = {
    'target_sample': 1000,
    'min_fresh_pct': 0.10,  # Very relaxed
    'max_panel_pct': 0.90,  # Very relaxed
    'tolerance': 0.20,       # Very loose
    'enforce_region_mins': False,  # Disable all
    'enforce_size_mins': False,
    'enforce_industry_mins': False,
}
```

### 3. Export for External Tools

```python
# Export data for manual analysis
df.to_excel('data_for_manual_analysis.xlsx')
```

---

**Last Updated:** January 2026
