### 82. **How do you create bidirectional lookups?**

Use CHOOSE with MATCH:
=INDEX(DataRange, MATCH(RowValue, RowHeaders, 0), MATCH(ColValue, ColHeaders, 0))

**Excel 365 alternative:**
Combine XLOOKUP or use FILTER with multiple conditions
