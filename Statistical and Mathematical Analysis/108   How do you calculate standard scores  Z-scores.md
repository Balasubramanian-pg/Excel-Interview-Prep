### 108. **How do you calculate standard scores (Z-scores)?**

=(Value - Mean) / StandardDeviation

Excel formula:
=(A1 - AVERAGE($A$1:$A$100)) / STDEV.S($A$1:$A$100)

**STANDARDIZE function:**
=STANDARDIZE(x, mean, standard_dev)
Example: =STANDARDIZE(A1, AVERAGE($A:$A), STDEV.S($A:$A))

**Use case:** Comparing values from different distributions
