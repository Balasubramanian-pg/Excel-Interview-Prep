### 32. **Explain all summing functions**

- **SUM(number1, number2, ...)**: Adds numbers
- **SUMIF(range, criteria, [sum_range])**: Sums based on one condition
Example: =SUMIF(A:A, "West", B:B) sums column B where column A is "West"
- **SUMIFS(sum_range, criteria_range1, criteria1, ...)**: Multiple criteria
Example: =SUMIFS(D:D, A:A, "West", B:B, ">1000")
- **SUBTOTAL(function_num, ref1, ...)**: Aggregate that ignores filtered rows
Function numbers: 1-11 (include hidden), 101-111 (exclude hidden)
Example: =SUBTOTAL(9, A1:A100) sums visible cells (9 = SUM)
