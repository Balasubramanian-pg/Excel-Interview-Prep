### 171. **Scientific: How do you calculate statistical significance?**

**Standard Error:**
=STDEV.S(Sample) / SQRT(COUNT(Sample))

**Confidence Interval:**

- Lower: =AVERAGE(Sample) - CONFIDENCE.T(0.05, STDEV.S(Sample), COUNT(Sample))
- Upper: =AVERAGE(Sample) + CONFIDENCE.T(0.05, STDEV.S(Sample), COUNT(Sample))

**T-Statistic:**
=(Sample_Mean - Population_Mean) / (Sample_StdDev / SQRT(Sample_Size))

**P-Value (from T-Test):**
=T.DIST.2T(ABS(T_Statistic), Degrees_of_Freedom)

**Effect Size (Cohen's d):**
=(Mean1 - Mean2) / Pooled_StdDev
