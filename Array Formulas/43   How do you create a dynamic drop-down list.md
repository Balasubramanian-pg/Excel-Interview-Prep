### 43. **How do you create a dynamic drop-down list?**

Use named range with OFFSET and COUNTA:
=OFFSET($A$1, 0, 0, COUNTA($A:$A), 1)

Or in Excel 365, simply reference the spilling array from UNIQUE or FILTER.
