### 92. **How do you use SORTBY with multiple criteria?**

=SORTBY(DataRange, SortBy1, Order1, SortBy2, Order2, ...)

Example:
=SORTBY(A1:C100, B1:B100, -1, C1:C100, 1)

Sorts data by column B descending, then column C ascending

**Dynamic sorted unique list:**
=SORT(UNIQUE(A1:A100))
