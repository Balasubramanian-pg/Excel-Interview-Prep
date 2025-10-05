### 110. **How do you calculate confidence intervals?**

**CONFIDENCE.NORM(alpha, standard_dev, size):**
Returns margin of error for confidence interval

Example: =CONFIDENCE.NORM(0.05, STDEV.S(A:A), COUNT(A:A))
95% confidence interval (alpha = 0.05)

**Full confidence interval:**

- Lower bound: =AVERAGE(A:A) - CONFIDENCE.NORM(0.05, STDEV.S(A:A), COUNT(A:A))
- Upper bound: =AVERAGE(A:A) + CONFIDENCE.NORM(0.05, STDEV.S(A:A), COUNT(A:A))
