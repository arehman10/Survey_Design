# Quick Start Tutorial

Get up and running with the Sample Allocation Optimizer in 10 minutes.

## Prerequisites

- Python 3.8+ installed
- Basic familiarity with command line
- Excel or CSV file with sample data

---

## Step 1: Installation (2 minutes)

### Download the Code

```bash
# Navigate to where you want to install
cd ~/Documents

# Extract the improved_solver folder here
```

### Install Dependencies

```bash
cd improved_solver
pip install -r requirements.txt
```

**Expected output:**
```
Successfully installed streamlit-1.28.0 pandas-2.0.0 cvxpy-1.4.0 ...
```

---

## Step 2: Prepare Your Data (3 minutes)

### Option A: Use Sample Data

We've provided sample data you can use immediately:

```bash
# Located at: examples/sample_data.csv
```

### Option B: Prepare Your Own Data

Create an Excel file with these columns:

| Column | Example |
|--------|---------|
| Region | North |
| Size | Small |
| Industry | Technology |
| Population | 1500 |
| PanelAvailable | 200 |

**Save as:** `my_data.xlsx`

### Quick Data Checklist

- ✅ All 5 columns present
- ✅ No missing values
- ✅ Numbers are numbers (no commas, $, etc.)
- ✅ Column names match exactly (case-sensitive)

---

## Step 3: Launch the App (1 minute)

```bash
cd src
streamlit run app.py
```

**Expected output:**
```
You can now view your Streamlit app in your browser.
Local URL: http://localhost:8501
```

Your browser should open automatically. If not, go to: http://localhost:8501

---

## Step 4: Run Your First Optimization (4 minutes)

### A. Upload Your Data

1. Click **"Browse files"**
2. Select your Excel file
3. Wait for "✅ Loaded X rows of data"
4. Expand **"Data Preview"** to verify

### B. Configure Scenario 1

In the left panel:

```
Target Sample Size: 1000
Min Fresh %: 20
Max Panel %: 80
Tolerance %: 5
```

Leave Advanced Constraints checked (recommended for first run).

### C. Configure Scenario 2

In the right panel:

```
Target Sample Size: 1500
Min Fresh %: 30
Max Panel %: 70
Tolerance %: 5
```

### D. Run Optimization

1. Click the big blue **"🚀 Run Optimization"** button
2. Wait for "✅ Scenario 1 optimized successfully"
3. Wait for "✅ Scenario 2 optimized successfully"

---

## Step 5: Review Results

### Summary Tab

Look for:
- ✅ Total Sample matches your target (±5%)
- ✅ Panel % within your limits
- ✅ Fresh % within your limits

**Example:**
```
Scenario 1:
Total Sample: 1,023
Panel Sample: 798 (78%)
Fresh Sample: 225 (22%)
```

### Detailed Results Tab

- Shows allocation for each Region × Size × Industry combination
- Check base weights (lower is better)
- Look for any 0 allocations

### Comparison Tab

- See differences between Scenario 1 and Scenario 2
- Useful for decision-making

---

## Step 6: Export Results

1. Scroll to bottom
2. Click **"Generate Excel Report"**
3. Click **"📥 Download Excel Report"**
4. Open the file in Excel

### What's in the Excel Report?

- **Summary**: Overview of both scenarios
- **Input_Data**: Your original data
- **S1_Detail, S2_Detail**: Detailed allocations
- **S1_Panel, S2_Panel**: Panel allocations
- **S1_Fresh, S2_Fresh**: Fresh allocations
- **Comparison**: Scenario differences

---

## Common First-Time Issues

### Issue: "Missing required columns"

**Fix:**
- Check column names: `Region` not `region`
- No extra spaces
- All 5 columns present

### Issue: "Solver status: infeasible"

**Fix:**
- Reduce target sample to 500
- Increase tolerance to 10%
- Uncheck "Enforce minimums" in Advanced

### Issue: Optimization takes too long

**Fix:**
- Use sample data first to test
- Reduce data size if yours is very large
- Check internet connection (not needed, but good practice)

---

## Next Steps

### Learn More

1. **Read the Full Guide:** `docs/README.md`
2. **API Reference:** `docs/API_REFERENCE.md`
3. **Troubleshooting:** `docs/TROUBLESHOOTING.md`

### Customize

1. **Adjust Parameters:**
   - Try different confidence levels
   - Experiment with panel/fresh ratios
   - Test various target sizes

2. **Multiple Scenarios:**
   - Compare conservative vs. aggressive
   - Test different constraint combinations

3. **Advanced Features:**
   - Save sessions for later
   - Generate HTML reports
   - Use as Python library

---

## Example Workflows

### Workflow 1: Conservative vs. Aggressive

**Scenario 1 (Conservative):**
```
Target: 1500
Min Fresh: 30%
Max Panel: 70%
Minimums: All enabled
```

**Scenario 2 (Aggressive):**
```
Target: 1000
Min Fresh: 15%
Max Panel: 85%
Minimums: Region only
```

**Compare:** Which gives better base weights?

### Workflow 2: Budget Constraints

**Scenario 1 (Full Budget):**
```
Target: 2000
Min Fresh: 20%
```

**Scenario 2 (Half Budget):**
```
Target: 1000
Min Fresh: 20%
```

**Compare:** Quality loss from budget cut?

### Workflow 3: Panel Utilization

**Scenario 1 (Max Panel):**
```
Target: 1500
Min Fresh: 10%
Max Panel: 90%
```

**Scenario 2 (Balanced):**
```
Target: 1500
Min Fresh: 40%
Max Panel: 60%
```

**Compare:** Panel vs. fresh trade-offs

---

## Tips for Success

### 1. Start Simple
- Use sample data first
- Default settings are good
- Add complexity gradually

### 2. Understand Your Data
- What's your total population?
- How much panel is available?
- Are segments balanced?

### 3. Set Realistic Targets
- Total sample < total panel (for panel-heavy)
- Consider minimum sample sizes
- Check statistical requirements

### 4. Iterate
- Run multiple scenarios
- Compare results
- Refine constraints

### 5. Validate Results
- Do base weights make sense?
- Are minimums met?
- Is distribution reasonable?

---

## Cheat Sheet

### Quick Commands

```bash
# Install
pip install -r requirements.txt

# Run app
cd src && streamlit run app.py

# Run tests
pytest tests/ -v

# Generate sample data
python examples/generate_sample_data.py
```

### Quick Config

```python
# Minimal config
config = {
    'target_sample': 1000,
    'min_fresh_pct': 0.20,
    'max_panel_pct': 0.80,
    'tolerance': 0.05,
}

# Typical config
config = {
    'target_sample': 1500,
    'min_fresh_pct': 0.25,
    'max_panel_pct': 0.75,
    'tolerance': 0.05,
    'confidence_level': 0.95,
    'margin_of_error': 0.05,
    'enforce_region_mins': True,
    'enforce_size_mins': True,
    'enforce_industry_mins': True,
}
```

### Quick Fixes

```python
# Problem: Infeasible
→ Increase tolerance: 0.10
→ Reduce target: 500
→ Disable minimums

# Problem: Slow
→ Use SCS solver
→ Increase iterations: 20000
→ Reduce data size

# Problem: High base weights
→ Increase target sample
→ Check minimums enabled
→ Review population distribution
```

---

## Success Checklist

After your first successful run, you should have:

- [x] Uploaded your data successfully
- [x] Configured two scenarios
- [x] Obtained "optimal" or "optimal_inaccurate" status
- [x] Reviewed summary statistics
- [x] Downloaded Excel report
- [x] Verified base weights are reasonable
- [x] Checked that constraints are satisfied

**Congratulations! You're now ready to use the Sample Allocation Optimizer effectively.**

---

## Getting Help

### Still Stuck?

1. Check `docs/TROUBLESHOOTING.md`
2. Review `docs/README.md`
3. Run tests: `pytest tests/ -v`
4. Try with sample data
5. Contact support

### Share Success

Found this helpful? Share your feedback!

---

**Last Updated:** January 2026
**Version:** 2.0.0
