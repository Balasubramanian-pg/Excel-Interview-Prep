### 182. **BI: How do you calculate variance analysis?**

**Absolute Variance:**
=Actual - Budget

**Percentage Variance:**
=(Actual - Budget) / Budget

**Favorable/Unfavorable Indicator:**
=IF(Category="Revenue",
IF(Actual>Budget, "Favorable", "Unfavorable"),
IF(Actual<Budget, "Favorable", "Unfavorable")
)

**Variance Explanation:**
=IFS(
ABS(Pct_Variance)<0.05, "Minimal",
ABS(Pct_Variance)<0.10, "Moderate",
ABS(Pct_Variance)<0.20, "Significant",
TRUE, "Critical"
)
