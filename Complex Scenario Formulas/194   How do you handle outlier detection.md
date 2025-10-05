### 194. **How do you handle outlier detection?**

**Z-Score Method:**
=ABS((Value - AVERAGE($A$1:$A$1000)) / STDEV.S($A$1:$A$1000))

**Outlier flag (Z-score > 3):**
=IF(ABS(Z_Score)>3, "Outlier", "Normal")

**IQR Method:**

- Q1: =QUARTILE.INC(Data, 1)
- Q3: =QUARTILE.INC(Data, 3)
- IQR: =Q3 - Q1
- Lower Bound: =Q1 - 1.5*IQR
- Upper Bound: =Q3 + 1.5*IQR
- Outlier: =IF(OR(Value<Lower_Bound, Value>Upper_Bound), "Outlier", "Normal")

**Modified Z-Score (more robust):**
=0.6745*(Value-MEDIAN($A$1:$A$1000))/MAD
Where MAD = Median Absolute Deviation
