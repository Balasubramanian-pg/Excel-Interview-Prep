### 34. **What is AGGREGATE function?**

Advanced version of SUBTOTAL with more functions and error handling:
Syntax: =AGGREGATE(function_num, options, array, [k])

Function numbers: 1-19 (including LARGE, SMALL, PERCENTILE, etc.)
Options control what to ignore: errors, hidden rows, nested subtotals

Example: =AGGREGATE(9, 6, A1:A100) sums while ignoring error values
