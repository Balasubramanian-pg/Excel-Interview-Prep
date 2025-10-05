### 218. **How do you use GROUPBY (when available)?**

**Note:** GROUPBY is being rolled out in Excel 365 preview

**Group and sum:**
=GROUPBY(Categories, Values, SUM)

**Multiple aggregations:**
=GROUPBY(Categories, Values, LAMBDA(vals,
HSTACK(SUM(vals), AVERAGE(vals), COUNT(vals))
))

**Grouped with conditions:**
=GROUPBY(
FILTER(Category, Amount>100),
FILTER(Amount, Amount>100),
SUM
)
