# Excel Summing Functions Explained

This guide provides a comprehensive overview of Excel's summing functions, from basic addition to conditional summing with single or multiple criteria, including specialized functions for filtered data.

## Function Syntax and Usage

### SUM Function
```
=SUM(number1, [number2], ...)
```
Adds all numbers in the specified arguments, including numbers, cell references, ranges, and constants.

**Parameters:**
- `number1, number2, ...`: Values to sum (1-255 arguments)
- Accepts: Numbers, cell references, ranges, arrays
- Ignores: Text, logical values, empty cells
- Handles: Negative numbers, decimals, percentages

### SUMIF Function
```
=SUMIF(range, criteria, [sum_range])
```
Adds cells specified by a given condition or criteria.

**Parameters:**
- `range`: The range to evaluate against the criteria
- `criteria`: The condition that determines which cells to sum
- `sum_range`: Optional range to sum (if omitted, sums the criteria range)
- Supports: Text matches, wildcards (*, ?), comparison operators

### SUMIFS Function
```
=SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2], [criteria2], ...)
```
Adds cells that meet multiple specified conditions (AND logic).

**Parameters:**
- `sum_range`: The range containing values to sum
- `criteria_range1, criteria_range2, ...`: Ranges to evaluate against criteria
- `criteria1, criteria2, ...`: Corresponding conditions for each range
- All conditions must be TRUE for a cell to be included in the sum

### SUBTOTAL Function
```
=SUBTOTAL(function_num, ref1, [ref2], ...)
```
Returns a subtotal using specified aggregation function, automatically ignoring rows hidden by filters.

**Parameters:**
- `function_num`: Number specifying the aggregation function
- `ref1, ref2, ...`: Ranges to include in the calculation
- Function numbers: 1-11 (include manually hidden rows), 101-111 (exclude manually hidden rows)

## Worked Examples

Given sample data:
```
A1: Region   B1: Product   C1: Sales   D1: Status
A2: East     B2: Apple     C2: 1000    D2: Active
A3: West     B3: Orange    C3: 1500    D3: Inactive
A4: East     B4: Banana    C4: 2000    D4: Active
A5: West     B5: Apple     C5: 1200    D5: Active
A6: North    B6: Orange    C6: 1800    D6: Active
```

**Basic summing functions:**
```
=SUM(C2:C6)
```
Returns: `7500` (sum of all Sales values)

```
=SUMIF(A2:A6, "East", C2:C6)
```
Returns: `3000` (sums Sales where Region is "East")

```
=SUMIFS(C2:C6, A2:A6, "West", D2:D6, "Active")
```
Returns: `1200` (sums Sales where Region = "West" AND Status = "Active")

**SUBTOTAL examples:**
```
=SUBTOTAL(9, C2:C6)
```
Returns: `7500` (sum of visible cells, function 9 = SUM)

If rows 3 and 5 are filtered out:
```
=SUBTOTAL(9, C2:C6)
```
Returns: `4800` (sum of only visible, unfiltered rows)

> [!NOTE]
> SUMIFS was introduced in Excel 2007 and has a different argument order than SUMIF. SUMIFS requires the sum_range first, while SUMIF has it as the third argument.

> [!IMPORTANT]
> SUBTOTAL function numbers 1-11 include values in manually hidden rows, while 101-111 exclude them. Both sets ignore rows hidden by filters.

> [!TIP]
> Use cell references for dynamic criteria in SUMIF/SUMIFS:
> `=SUMIF(A:A, E1, C:C)` where E1 contains the criteria text
> `=SUMIFS(C:C, A:A, ">"&F1, B:B, G1)` for numeric comparisons

## Alternative Methods

### Using SUMPRODUCT for Complex Logic
For OR conditions or mathematical operations:
```
=SUMPRODUCT((A2:A6="East")*(C2:C6))
```
Sums Sales where Region is "East" (similar to SUMIF)

```
=SUMPRODUCT(((A2:A6="East")+(A2:A6="West"))*(C2:C6))
```
Sums Sales where Region is "East" OR "West"

### Using AGGREGATE Function
For advanced filtering and error handling:
```
=AGGREGATE(9, 5, C2:C6)
```
Sums range while ignoring error values and hidden rows (9 = SUM, 5 = ignore hidden rows and errors)

### Using Database Functions
For structured data analysis:
```
=DSUM(A1:D6, "Sales", F1:G2)
```
Where F1:G2 contains criteria range with field headers and conditions

### SUBTOTAL Function Numbers
Common function_num values:
- `9` or `109`: SUM
- `1` or `101`: AVERAGE
- `2` or `102`: COUNT
- `3` or `103`: COUNTA
- `4` or `104`: MAX
- `5` or `105`: MIN
