<script type="text/javascript" async
  src="https://cdnjs.cloudflare.com/ajax/libs/mathjax/2.7.5/MathJax.js?config=TeX-MML-AM_CHTML">
</script>


# 1. Pen-and-Paper Formulation

## 1.1 Decision Variables

Let there be $$\(n\)$$ cells. Each cell $$\(i\)$$ (for $$\(i = 1,\dots,n\))$$ corresponds to a combination of Region, Size, and Industry in the data. The decision variable is:

$$
x_i \quad (\text{the sample size chosen for cell } i).
$$

## 1.2 Objective Function

The code seeks to minimize the deviation between $$\(x_i\)$$ and a “proportional” target $$\(p_i\)$$.

$$
p_i = \text{PropSample}_i
$$

is computed as:

$$
p_i = \text{Population}_i \times \frac{\text{totalsample}}{\sum_{j=1}^n \text{Population}_j}
$$
The objective is:

$$
\min \sum_{i=1}^{n} \bigl(x_i - p_i\bigr)^2.
$$

This is a **least-squares** type objective that penalizes deviation of $$\(x_i\)$$ from the proportional target $$\(p_i\)$$.

## 1.3 Constraints

### 1.3.1 Sum of Samples Must Equal Total

$$
\sum_{i=1}^{n} x_i \;=\; \text{total sample}.
$$

### 1.3.2 Cell-Level Lower and Upper Bounds

Each cell $$\(i\)$$ has:

- **A lower bound**:

  $$
  x_i \;\ge\; \max\!\Bigl(
    \tfrac{\text{Population}_i}{\text{max base weight}},
    \;\text{min cell size},
    \;0
  \Bigr).
  $$

  This ensures each cell has enough sample to limit base weights and respect the cell’s minimum size requirement.

- **An upper bound**:

  $$
  x_i \;\le\; \min\!\Bigl(
    \text{Population}_i,
    \;\text{max cell size},
    \;\lceil \text{Population}_i \times \text{conversion rate} \rceil
  \Bigr).
  $$

  This ensures the sample drawn for a cell does not exceed realistic or user-imposed limits.

### 1.3.3 Dimension-Wise Minimums

For a given dimension (e.g., Region, Size, or Industry), let’s say for Region $$\(r\)$$, we require:

$$
\sum_{i \in \,\text{cells for region }r} x_i \;\ge\;\text{dimensionmins}[\text{Region}][r].
$$

The code generalizes this to any specified dimension (Region, Size, or Industry).

### 1.3.4 Integrality

$$
x_i \;\in\; \mathbb{Z}_{\ge 0}
\quad (\text{each } x_i \text{ is a non-negative integer}).
$$

## 1.4 Putting It All Together

In mathematical form:

$$
\begin{aligned}
&\text{Minimize} 
  && \sum_{i=1}^{n} \bigl(x_i - p_i\bigr)^2 \\[6pt]
&\text{subject to} \\[-1pt]
& \sum_{i=1}^{n} x_i = \text{total\_sample}, \\[6pt]
& x_i \;\ge\; \max\!\Bigl(
    \tfrac{\text{Pop}_i}{\text{max base weight}},
    \;\text{min cell size},
    \;0
  \Bigr), 
  && \quad \forall i, \\[6pt]
& x_i \;\le\; \min\!\Bigl(
    \text{Pop}_i,
    \;\text{max\_cell\_size},
    \;\bigl\lceil \text{Pop}_i \times \text{conversion\_rate}\bigr\rceil
  \Bigr),
  && \quad \forall i, \\[6pt]
& \sum_{i \in D(r)} x_i \;\ge\; \text{dimension\_mins}[D][r],
  && \quad \forall \text{dimension }D,\;\forall \text{value }r,\\[6pt]
& x_i \;\in\; \mathbb{Z}_{\ge 0},
  && \quad \forall i.
\end{aligned}
$$

---

# 2. Closed-Form Solution (Under Simplified Assumptions)

In general, because $$\(x_i\)$$ must be integer and must satisfy multiple lower/upper bounds and dimension constraints, there is **no simple closed-form formula** for the exact solution.

However, if we **ignore**:

1. Integrality (allowing $$\(x_i\)$$ to be any real number),
2. The lower bound constraints,
3. The upper bound constraints,
4. The dimension minimum constraints,

then the problem reduces to:

$$
\min \sum_{i=1}^{n} (x_i - p_i)^2 
\quad \text{subject to} \quad
\sum_{i=1}^{n} x_i = \sum_{i=1}^{n} p_i.
$$

Since 
$$\sum_{i=1}^n p_i = \text{total sample},$$ 
the constraint 
$$\sum_{i=1}^n x_i = \sum_{i=1}^n p_i$$ 
is automatically satisfied by choosing \(x_i = p_i\). That is the global minimum, giving an objective of 0.

In that simplified scenario, the closed-form optimal solution is:

$$
x_i^* = p_i.
$$

Once we impose the integer requirement (plus additional bounds), we must use a **mixed-integer optimization** method (like the one in the code) instead of a simple formula.

---

# 3. Step-by-Step Code Explanation

### Data Reshaping
The code transforms your wide-format data (e.g., columns for each Industry, plus Region/Size identifiers) into a long format where each row is a \((\text{Region}, \text{Size}, \text{Industry}, \text{Population})\) tuple.

### Proportional Targets
For each cell $$\(i\)$$, it computes the proportional target $$\(p_i\)$$ based on population shares and `total_sample`.

### Feasibility Checks
The function **`detailed_feasibility_check`** verifies if it's possible to meet:

- Minimum cell sizes
- Maximum cell sizes
- Maximum base weight constraints
- Dimension-wise minimums
- Conversion rate limits

### Formulating the MIP
1. Declares integer decision variables $$\(x_i\)$$.
2. Minimizes $$\(\sum (x_i - p_i)^2\)$$.
3. Constrains $$\(\sum x_i = \text{total sample}\)$$.
4. Imposes lower and upper bounds for each cell.
5. Applies dimension-wise minimums if given.

### Solving
- Tries solvers like **SCIP** or **ECOS\_BB** via CVXPY.
- If no feasible solution is found, a slack-based diagnostic helps identify which constraints are violated.

### Outputs
On success, returns the integer $$\(x_i\)$$ and computes the corresponding base weight $$\(\tfrac{\text{Population}_i}{x_i}\) (if \(x_i > 0\))$$.

---

# 4. Brief Description of the Solver

The code uses [CVXPY](https://www.cvxpy.org/) to model the **mixed-integer quadratic problem**. CVXPY’s role:

- **Model Construction**: Define the decision variables, the objective function (least squares), and the constraints (equalities, inequalities, integrality).
- **Solver Backend**: 
  - **SCIP**: A well-known solver for mixed-integer optimization problems.  
  - **ECOS\_BB**: A branch-and-bound variant of the ECOS solver that supports integer constraints.
- **Result**: The solver attempts to find a feasible solution that minimizes $$\(\sum (x_i - p_i)^2\)$$. If infeasible, the code can diagnose which constraints cause the conflict.

**Happy Optimizing!**
