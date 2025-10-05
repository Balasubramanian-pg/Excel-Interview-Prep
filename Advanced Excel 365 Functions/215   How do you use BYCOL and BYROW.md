### 215. **How do you use BYCOL and BYROW?**

**Column-wise sum:**
=BYCOL(A1:E10, LAMBDA(col, SUM(col)))

**Row-wise maximum:**
=BYROW(A1:E10, LAMBDA(row, MAX(row)))

**Column-wise average excluding outliers:**
=BYCOL(Data, LAMBDA(col,
AVERAGE(FILTER(col, ABS(col-AVERAGE(col))<2*STDEV(col)))
))

**Row-wise concatenation:**
=BYROW(A1:C10, LAMBDA(row, TEXTJOIN(", ", TRUE, row)))

**Complex aggregation by row:**
=BYROW(Sales_Data, LAMBDA(row,
INDEX(row,1) * INDEX(row,2) * (1-INDEX(row,3))
))
