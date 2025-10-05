### 51. **What is OFFSET and when do you use it?**

Returns reference offset from starting cell:
Syntax: =OFFSET(reference, rows, cols, [height], [width])

Examples:

- =OFFSET(A1, 2, 3) references cell D3 (2 rows down, 3 columns right)
- =SUM(OFFSET(A1, 0, 0, 10, 1)) sums 10 cells starting from A1
- =OFFSET(A1, 0, 0, COUNTA(A:A), 1) dynamic range expanding with data

**Use cases:**

- Dynamic named ranges
- Moving averages
- Creating flexible ranges

**Warning:** OFFSET is volatile, use with caution on large datasets
