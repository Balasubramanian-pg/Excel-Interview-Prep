### 219. **How do you use PIVOTBY (when available)?**

**Create pivot-like structure:**
=PIVOTBY(Row_Values, Column_Values, Data_Values, SUM)

**Multiple value fields:**
=PIVOTBY(Rows, Cols, Data, LAMBDA(vals,
HSTACK(SUM(vals), AVERAGE(vals))
))

**With grand totals:**
=VSTACK(
HSTACK("", Unique_Cols, "Total"),
HSTACK(Unique_Rows, Pivot_Data, Row_Totals),
HSTACK("Total", Col_Totals, Grand_Total)
)
