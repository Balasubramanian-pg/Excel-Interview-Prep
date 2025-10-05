### 97. **What is EXPAND and CHOOSEROWS/CHOOSECOLS?**

**EXPAND(array, rows, [cols], [pad_with]):**
Expands array to specified size:
=EXPAND(A1:B5, 10, 3, "N/A")

**CHOOSEROWS(array, row_num1, ...):**
=CHOOSEROWS(A1:C100, 1, 5, 10)
Returns rows 1, 5, and 10

**CHOOSECOLS(array, col_num1, ...):**
=CHOOSECOLS(A1:E100, 1, 3, 5)
Returns columns 1, 3, and 5
