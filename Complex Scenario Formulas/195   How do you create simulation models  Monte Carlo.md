### 195. **How do you create simulation models (Monte Carlo)?**

**Random scenario generator:**
=Mean + NORM.INV(RAND(), 0, 1) * StdDev

**Multiple correlated variables:**
Requires Cholesky decomposition (complex, typically use add-ins)

**Simple profit simulation:**
=RANDARRAY(1000, 1, Min_Revenue, Max_Revenue) - RANDARRAY(1000, 1, Min_Cost, Max_Cost)

**Probability of success:**
=COUNTIF(Simulation_Results, ">0") / 1000
