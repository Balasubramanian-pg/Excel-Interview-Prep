### 52. **How do you create a dynamic named range?**

Formula Manager → New Name:
=OFFSET(Sheet1!$A$1, 0, 0, COUNTA(Sheet1!$A:$A), 1)

This creates a range that automatically expands/contracts with data in column A.

**Excel 365 alternative:**
Simply name a cell with a FILTER or spilling formula, and the name automatically includes the spilled range.
