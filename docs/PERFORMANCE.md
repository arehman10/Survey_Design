# Performance Optimization Guide

Techniques for improving the speed and efficiency of the Sample Allocation Optimizer.

## Table of Contents

1. [Understanding Performance](#understanding-performance)
2. [Data Optimization](#data-optimization)
3. [Solver Configuration](#solver-configuration)
4. [Constraint Optimization](#constraint-optimization)
5. [Code-Level Optimizations](#code-level-optimizations)
6. [Benchmarking](#benchmarking)

---

## Understanding Performance

### Performance Metrics

The optimizer's performance depends on:

1. **Problem Size:**
   - Number of cells (Region × Size × Industry combinations)
   - Total population
   - Number of constraints

2. **Solver Complexity:**
   - Which solver is used
   - Number of iterations needed
   - Numerical conditioning

3. **Data Quality:**
   - Sparsity
   - Scale differences
   - Constraint tightness

### Typical Performance

| Problem Size | Cells | Time (ECOS) | Time (SCS) |
|--------------|-------|-------------|------------|
| Small        | <50   | <1s         | <1s        |
| Medium       | 50-200| 1-5s        | 1-3s       |
| Large        | 200-500| 5-30s      | 3-15s      |
| Very Large   | >500  | 30s-5min    | 10s-1min   |

---

## Data Optimization

### 1. Reduce Data Dimensionality

**Problem:** Too many small segments create unnecessary complexity

**Solution:** Aggregate small segments

```python
# Remove very small segments
min_population = 50
df = df[df['Population'] >= min_population]

# Combine rare industries
industry_counts = df['Industry'].value_counts()
rare_industries = industry_counts[industry_counts < 3].index
df.loc[df['Industry'].isin(rare_industries), 'Industry'] = 'Other'

# Result: Fewer cells, faster optimization
```

### 2. Balance Data

**Problem:** Extreme imbalances in population or panel

**Solution:** Log transformation or standardization

```python
# For very large population ranges
df['PopulationLog'] = np.log1p(df['Population'])

# Or standardize
from sklearn.preprocessing import StandardScaler
scaler = StandardScaler()
df['PopulationScaled'] = scaler.fit_transform(df[['Population']])
```

### 3. Pre-filter Infeasible Cells

**Problem:** Cells with no panel available still create variables

**Solution:** Remove or mark before optimization

```python
# Remove cells where panel is 0 and fresh is required
if config.get('allow_zero_sample', False):
    pass  # Keep all
else:
    # Remove cells that can't contribute
    df = df[(df['PanelAvailable'] > 0) | (df['Population'] > 0)]
```

---

## Solver Configuration

### 1. Choose the Right Solver

**ECOS (Default):**
- Best for: Medium problems with strict integer constraints
- Pros: Accurate, handles integers well
- Cons: Slower for large problems

```python
config['solver'] = 'ECOS'
config['max_iterations'] = 10000
```

**SCS:**
- Best for: Large problems, looser tolerances OK
- Pros: Fast, scales well
- Cons: Less accurate, may need rounding

```python
config['solver'] = 'SCS'
config['max_iterations'] = 5000  # Usually converges faster
```

**GLPK_MI:**
- Best for: Strict integer requirements
- Pros: Handles pure integer programs
- Cons: Slowest for large problems

```python
config['solver'] = 'GLPK_MI'
```

### 2. Tune Solver Parameters

**For ECOS:**

```python
from cvxpy import ECOS

problem.solve(
    solver=ECOS,
    max_iters=10000,
    abstol=1e-7,    # Absolute tolerance
    reltol=1e-6,    # Relative tolerance
    feastol=1e-7,   # Feasibility tolerance
    verbose=False   # Disable logging for speed
)
```

**For SCS:**

```python
from cvxpy import SCS

problem.solve(
    solver=SCS,
    max_iters=5000,
    eps=1e-4,       # Looser for speed
    alpha=1.5,      # Step size
    verbose=False
)
```

### 3. Adaptive Solving

Try fast solver first, fall back if needed:

```python
def adaptive_solve(problem, config):
    """Try SCS first, fall back to ECOS if needed."""
    
    # Try fast solver
    try:
        problem.solve(solver='SCS', max_iters=3000)
        if problem.status in ['optimal', 'optimal_inaccurate']:
            return problem.status
    except:
        pass
    
    # Fall back to accurate solver
    problem.solve(solver='ECOS', max_iters=10000)
    return problem.status
```

---

## Constraint Optimization

### 1. Relax Unnecessary Constraints

**Problem:** Too many constraints slow solving

**Solution:** Only enforce essential constraints

```python
# Minimal constraints (fastest)
config = {
    'enforce_region_mins': False,
    'enforce_size_mins': False,
    'enforce_industry_mins': False,
    'tolerance': 0.15  # Wider tolerance
}

# Vs. full constraints (slower but more accurate)
config = {
    'enforce_region_mins': True,
    'enforce_size_mins': True,
    'enforce_industry_mins': True,
    'tolerance': 0.05
}
```

### 2. Use Soft Constraints

**Problem:** Hard constraints may be infeasible

**Solution:** Convert to penalties in objective

```python
# Instead of hard constraint:
# total_sample >= target

# Use soft constraint:
objective = cp.Minimize(
    cp.square(cp.sum(total_sample) - target) +  # Penalty for deviation
    regularization_term
)
```

### 3. Simplify Minimum Calculations

**Problem:** Complex minimum calculations

**Solution:** Use simpler formulas for large N

```python
def fast_minimum(N, confidence=0.95, moe=0.05):
    """Fast approximation for large N."""
    if N > 10000:
        # For large N, FPC ≈ n_infinity
        z = 1.96 if confidence == 0.95 else 1.645
        return (z**2) * 0.25 / (moe**2)
    else:
        # Use full formula
        return compute_fpc_min(N, n_infinity)
```

---

## Code-Level Optimizations

### 1. Vectorize Operations

**Problem:** Loops are slow

**Solution:** Use NumPy vectorization

```python
# Slow (loop)
base_weights = []
for i in range(len(df)):
    if df.loc[i, 'Sample'] > 0:
        bw = df.loc[i, 'Population'] / df.loc[i, 'Sample']
    else:
        bw = 0
    base_weights.append(bw)

# Fast (vectorized)
base_weights = np.where(
    df['Sample'] > 0,
    df['Population'] / df['Sample'],
    0
)
```

### 2. Use Efficient Data Structures

```python
# Slow: Repeated filtering
for region in regions:
    df_region = df[df['Region'] == region]
    # process...

# Fast: Group once
for region, df_region in df.groupby('Region'):
    # process...
```

### 3. Cache Repeated Calculations

```python
class OptimizedOptimizer:
    def __init__(self):
        self._z_score_cache = {}
    
    def get_z_score(self, confidence):
        if confidence not in self._z_score_cache:
            self._z_score_cache[confidence] = self._calculate_z_score(confidence)
        return self._z_score_cache[confidence]
```

### 4. Parallel Processing (Advanced)

For multiple scenarios:

```python
from concurrent.futures import ProcessPoolExecutor

def optimize_scenario(data, config):
    optimizer = SampleOptimizer()
    return optimizer.optimize(data, config)

# Run scenarios in parallel
with ProcessPoolExecutor(max_workers=2) as executor:
    futures = [
        executor.submit(optimize_scenario, data, config1),
        executor.submit(optimize_scenario, data, config2)
    ]
    results = [f.result() for f in futures]
```

---

## Benchmarking

### 1. Profile Your Code

```python
import cProfile
import pstats

# Profile optimization
profiler = cProfile.Profile()
profiler.enable()

result = optimizer.optimize(data, config)

profiler.disable()
stats = pstats.Stats(profiler)
stats.sort_stats('cumulative')
stats.print_stats(20)  # Top 20 functions
```

### 2. Measure Solve Time

```python
import time

def benchmark_solver(data, config, n_runs=5):
    """Benchmark solver performance."""
    times = []
    
    for i in range(n_runs):
        start = time.time()
        result = optimizer.optimize(data, config)
        elapsed = time.time() - start
        
        times.append(elapsed)
        print(f"Run {i+1}: {elapsed:.2f}s - Status: {result.get('status')}")
    
    print(f"\nAverage: {np.mean(times):.2f}s")
    print(f"Std Dev: {np.std(times):.2f}s")
    return times
```

### 3. Compare Solvers

```python
def compare_solvers(data, config):
    """Compare performance of different solvers."""
    solvers = ['ECOS', 'SCS', 'GLPK_MI']
    results = {}
    
    for solver in solvers:
        config['solver'] = solver
        start = time.time()
        
        try:
            result = optimizer.optimize(data, config)
            elapsed = time.time() - start
            
            results[solver] = {
                'time': elapsed,
                'status': result.get('status'),
                'objective': result.get('objective_value')
            }
        except Exception as e:
            results[solver] = {'error': str(e)}
    
    return pd.DataFrame(results).T
```

---

## Optimization Checklist

Before optimizing, check:

- [ ] Is the problem actually slow? (Measure first)
- [ ] Is the data clean and minimal?
- [ ] Are constraints necessary?
- [ ] Is the right solver being used?
- [ ] Are solver parameters tuned?
- [ ] Can constraints be relaxed?
- [ ] Are there obvious bottlenecks?

---

## Performance Targets

### Good Performance

- **Small problems** (<100 cells): <2s
- **Medium problems** (100-300 cells): <10s
- **Large problems** (300-1000 cells): <60s

### If Slower

1. **Check data size:**
```python
n_cells = len(df)
print(f"Cells: {n_cells}")
if n_cells > 500:
    print("Consider aggregating data")
```

2. **Check constraints:**
```python
n_constraints = len(problem.constraints)
print(f"Constraints: {n_constraints}")
if n_constraints > 1000:
    print("Consider relaxing constraints")
```

3. **Try different solver:**
```python
config['solver'] = 'SCS'  # Usually faster
```

---

## Advanced Techniques

### 1. Warm Start

Use previous solution as starting point:

```python
# Save previous solution
prev_panel = result['df_long']['PanelSample'].values
prev_fresh = result['df_long']['FreshSample'].values

# Use as warm start (solver-dependent)
panel.value = prev_panel
fresh.value = prev_fresh

problem.solve(warm_start=True)
```

### 2. Problem Decomposition

Break large problem into smaller sub-problems:

```python
def decompose_by_region(data, config):
    """Solve each region separately, then combine."""
    results = []
    
    for region in data['Region'].unique():
        df_region = data[data['Region'] == region]
        
        # Adjust target proportionally
        region_config = config.copy()
        region_config['target_sample'] = int(
            config['target_sample'] * 
            df_region['Population'].sum() / data['Population'].sum()
        )
        
        result = optimizer.optimize(df_region, region_config)
        results.append(result['df_long'])
    
    return pd.concat(results, ignore_index=True)
```

### 3. Hierarchical Optimization

Optimize in stages:

```python
def hierarchical_optimize(data, config):
    """Optimize in stages: Region → Size → Industry."""
    
    # Stage 1: Allocate to regions
    region_allocation = allocate_to_regions(data, config)
    
    # Stage 2: Within each region, allocate to sizes
    size_allocation = allocate_to_sizes(region_allocation, config)
    
    # Stage 3: Within each size, allocate to industries
    final_allocation = allocate_to_industries(size_allocation, config)
    
    return final_allocation
```

---

## Monitoring Performance

### Real-Time Monitoring

```python
import logging

logging.basicConfig(level=logging.INFO)

# In optimizer
logger.info(f"Building problem: {n_variables} variables, {n_constraints} constraints")
logger.info(f"Solving with {solver}...")

start = time.time()
problem.solve()
logger.info(f"Solved in {time.time() - start:.2f}s")
```

### Performance Metrics Dashboard

```python
def create_performance_report(result):
    """Generate performance metrics."""
    metrics = {
        'Total Time': result.get('solve_time', 0),
        'Status': result.get('status'),
        'Iterations': result.get('iterations'),
        'Objective Value': result.get('objective_value'),
        'Variables': len(result.get('df_long', [])),
        'Constraints': '(see problem)'
    }
    
    return pd.DataFrame([metrics]).T
```

---

## Summary

**Key Takeaways:**

1. **Start with data** - Reduce, clean, balance
2. **Choose right solver** - SCS for speed, ECOS for accuracy
3. **Relax constraints** - Only enforce what's necessary
4. **Measure first** - Profile before optimizing
5. **Iterate** - Small improvements add up

**Quick Wins:**

- Use SCS solver: ~2-3x speedup
- Increase tolerance to 0.10: ~1.5x speedup
- Disable unnecessary minimums: ~1.5-2x speedup
- Remove small segments: Varies by data

**Combined effect:** Often 5-10x faster with minimal quality loss

---

**Last Updated:** January 2026
