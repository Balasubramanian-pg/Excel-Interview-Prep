### 100. **How do you create dynamic chart ranges?**

Named range formula:
=OFFSET(Sheet1!$A$1, 0, 0, COUNTA(Sheet1!$A:$A), COUNTA(Sheet1!$1:$1))

Creates range that expands with both rows and columns

**Excel 365:** Simply use table or dynamic array formula
