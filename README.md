# Explaination of the Survey Design App

## 1.1 Decision Variables

Let there be \(n\) cells. Each cell \(i\) (for \(i = 1,\dots,n\)) corresponds to a combination of Region, Size, and Industry in the data. The decision variable is:
'''
\[
x_i \quad \text{(the sample size chosen for cell } i\text{)}.
\]
'''
## 1.2 Objective Function

The code seeks to minimize the deviation between \(x_i\) and a “proportional” target \(p_i\). 

\[
p_i = \text{PropSample}_i
\]
is computed as:  
\[
p_i 
= \text{Population}_i 
  \times \frac{\text{total\_sample}}
             {\sum_{j=1}^{n} \text{Population}_j}.
\]

The objective is:
\[
\min \sum_{i=1}^{n} (x_i - p_i)^2.
\]
This is a **least-squares** type objective that penalizes deviation of \(x_i\) from the proportional target \(p_i\).

## 1.3 Constraints

### 1.3.1 Sum of Samples Must Equal Total
\[
\sum_{i=1}^{n} x_i = \text{total\_sample}.
\]

### 1.3.2 Cell-Level Lower and Upper Bounds

Each cell \(i\) has:

- **A lower bound**:
  \[
  x_i \;\ge\; \max\Bigl(\frac{\text{Population}_i}{\text{max\_base\_weight}},\;\text{min\_cell\_size},\;0\Bigr).
  \]
  This ensures each cell has enough sample to limit base weights and respect the cell’s minimum size requirement.

- **An upper bound**:
  \[
  x_i \;\le\; \min\Bigl(\text{Population}_i,\;\text{max\_cell\_size},\;\lceil \text{Population}_i \times \text{conversion\_rate}\rceil\Bigr).
  \]
  This ensures the sample drawn for a cell does not exceed realistic or user-imposed limits.

### 1.3.3 Dimension-Wise Minimums

For a given dimension (e.g., Region, Size, or Industry), let’s say for Region \(r\), we require:
\[
\sum_{i \in\,\text{cells for region }r} x_i \;\ge\;\text{dimension\_mins}[\text{Region}][r].
\]
The code generalizes this to any specified dimension (Region, Size, or Industry).

### 1.3.4 Integrality

\[
x_i \in \mathbb{Z}_{\ge 0}\quad (\text{each } x_i \text{ is a non-negative integer}).
\]

## 1.4 Putting It All Together

In mathematical form:

\[
\begin{aligned}
&\text{Minimize} && \sum_{i=1}^{n} (x_i - p_i)^2 \\[6pt]
&\text{subject to} \\[-1pt]
& \sum_{i=1}^{n} x_i = \text{total\_sample}, \\[6pt]
& x_i \;\ge\; \max\left(\frac{\text{Pop}_i}{\text{max\_base\_weight}},\;\text{min\_cell\_size},\;0\right), \quad \forall i, \\[6pt]
& x_i \;\le\; \min\left(\text{Pop}_i,\;\text{max\_cell\_size},\;\left\lceil \text{Pop}_i \times \text{conversion\_rate}\right\rceil\right), \quad \forall i, \\[6pt]
& \sum_{i \in D(r)} x_i \;\ge\; \text{dimension\_mins}[D][r], \quad \forall \text{dimension }D,\;\forall \text{value }r,\\[6pt]
& x_i \in \mathbb{Z}_{\ge 0}, \quad \forall i.
\end{aligned}
\]

---

# 2. Closed-Form Solution (Under Simplified Assumptions)

In general, because \(x_i\) must be integer and must satisfy multiple lower/upper bounds and dimension constraints, there is **no simple closed-form formula** for the exact solution.

However, if we **ignore**:

1. Integrality (allowing \(x_i\) to be any real number),
2. The lower bound constraints,
3. The upper bound constraints,
4. The dimension minimum constraints,

then the problem reduces to:

\[
\min \sum_{i=1}^{n} (x_i - p_i)^2 
\quad \text{subject to} \quad
\sum_{i=1}^{n} x_i = \sum_{i=1}^{n} p_i.
\]

Since \(\sum_{i=1}^n p_i = \text{total\_sample}\), the constraint \(\sum_{i=1}^n x_i = \sum_{i=1}^n p_i\) is automatically satisfied by choosing \(x_i = p_i\). That is the global minimum, giving an objective of 0.

In that simplified scenario, the closed-form optimal solution is:

\[
x_i^* = p_i.
\]

Once we impose the integer requirement (plus additional bounds), we must use a **mixed-integer optimization method** (like the one in the code) instead of a simple formula.

---

# 3. Step-by-Step Code Explanation

Below is a high-level explanation of how the code works (referring to the main solver function and related helpers):

1. **Data Reshaping**  
   - The code transforms your wide-format data (e.g., columns for each Industry, plus Region/Size identifiers) into a long format where each row is a (Region, Size, Industry) tuple with an associated population.

2. **Proportional Targets**  
   - For each cell \(i\), it computes a proportional target \(p_i\) based on population shares and the desired total sample size.

3. **Feasibility Checks**  
   - A function (`detailed_feasibility_check`) verifies if it’s even possible to meet certain requirements:
     - Minimum cell sizes,
     - Maximum cell sizes,
     - Maximum base weight constraints,
     - Dimension-wise minimums,
     - Conversion rate limits.
   - If these constraints are collectively impossible, it reports the conflicts.

4. **Formulating the MIP**  
   - Defines integer decision variables \(x_i\) for each cell.
   - Minimizes \(\sum (x_i - p_i)^2\).
   - Constrains:
     - \(\sum x_i = \text{total\_sample}\).
     - Lower bounds for each cell (based on min cell sizes and base weight).
     - Upper bounds for each cell (based on max cell sizes, population, and conversion rate).
     - Dimension-wise minimum constraints.

5. **Solving**  
   - Attempts to solve the Mixed-Integer Quadratic Problem (MIQP) via CVXPY.
   - Tries solvers like **SCIP** or **ECOS_BB**. If they fail, it diagnoses infeasibility using slack variables.

6. **Generating Outputs**  
   - Once a solution \(x_i\) is found, it calculates the **base weight** as \(\frac{\text{Population}_i}{x_i}\) when \(x_i > 0\), and merges the results back into a final table.

---

# 4. Brief Description of the Solver

The code uses the [**CVXPY**](https://www.cvxpy.org/) library to model and solve a **mixed-integer optimization** problem. Key points:

- **Modeling:**  
  You declare decision variables (`cp.Variable(n_cells, integer=True)`), define an objective function (`cp.sum_squares(x - df_long["PropSample"])`), and specify constraints (equalities, inequalities, integrality).

- **Solvers:**  
  - **SCIP** ([link](https://www.scipopt.org/)) is a well-known mixed-integer programming solver.
  - **ECOS_BB** ([link](https://github.com/embotech/ecos)) is a branch-and-bound extension of ECOS for mixed-integer conic optimization.  
  CVXPY automatically translates your problem into the appropriate format for the chosen solver.

- **Outcome:**  
  The solver searches for a feasible \(x\) that satisfies all constraints and **minimizes** the least-squares objective. If no feasible solution exists, the code can diagnose which constraints caused the infeasibility.

---

**Happy Optimizing!**
