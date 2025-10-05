### 190. **How do you handle missing data imputation?**

**Forward Fill:**
=IF(ISBLANK(A2), B1, A2)

**Backward Fill:**
=IF(ISBLANK(A2), A3, A2)

**Linear Interpolation:**
=IF(ISBLANK(B2),
Previous_Value + ((Next_Value-Previous_Value)/(Next_Date-Previous_Date))*(B2_Date-Previous_Date),
B2
)

**Mean Imputation:**
=IF(ISBLANK(A2), AVERAGE($A$2:$A$1000), A2)

**Excel 365 - Remove blanks:**
=FILTER(A:A, A:A<>"")
