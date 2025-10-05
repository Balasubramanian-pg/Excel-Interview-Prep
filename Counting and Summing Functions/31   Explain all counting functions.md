### 31. **Explain all counting functions**

- **COUNT(value1, value2, ...)**: Counts cells with numbers
- **COUNTA(value1, value2, ...)**: Counts non-empty cells
- **COUNTBLANK(range)**: Counts empty cells
- **COUNTIF(range, criteria)**: Counts cells meeting one condition
Example: =COUNTIF(A1:A100, ">50")
- **COUNTIFS(range1, criteria1, range2, criteria2, ...)**: Multiple criteria
Example: =COUNTIFS(A:A, "West", B:B, ">1000", C:C, "<5000")
