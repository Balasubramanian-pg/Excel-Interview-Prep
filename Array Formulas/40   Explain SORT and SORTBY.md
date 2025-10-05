### 40. **Explain SORT and SORTBY**

**SORT(array, [sort_index], [sort_order], [by_col]):**
Sorts array by column/row
Example: =SORT(A1:C100, 2, -1) sorts by column 2 descending

**SORTBY(array, by_array1, [order1], ...):**
Sorts by different criteria
Example: =SORTBY(A1:C100, B1:B100, -1, C1:C100, 1)
Sorts by column B descending, then column C ascending
