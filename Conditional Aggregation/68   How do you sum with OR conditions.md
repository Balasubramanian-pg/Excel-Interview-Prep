# How to Sum with OR Conditions in Excel

This guide explains multiple methods to sum values based on OR logic, where values are summed if they meet at least one of several specified conditions.

## Formula Syntax

### Method 1: Multiple SUMIF Functions
```
=SUMIF(range, criteria1, sum_range) + SUMIF(range, criteria2, sum_range)
```

**Parameters:**
- `range`: The range to evaluate for each condition
- `criteria1`, `criteria2`: Individual conditions to check
- `sum_range`: The range containing values to sum

### Method 2: SUMPRODUCT with Addition
```
=SUMPRODUCT(((range=criteria1) + (range=criteria2)), sum_range)
```

**Parameters:**
- `range=criteria1`: Boolean array for first condition
- `range=criteria2`: Boolean array for second condition
- The `+` operator implements OR logic
- `sum_range`: Values to sum

### Method 3: Array Formula with IF
```
=SUM(IF((range=criteria1) + (range=criteria2), sum_range, 0))
```

**Parameters:**
- Requires Ctrl+Shift+Enter in older Excel versions
- `IF` function returns values where either condition is TRUE

## Worked Example

Given the following data:
```
A1: Region    B1: Sales
A2: West      B2: 1000
A3: East      B3: 1500
A4: North     B4: 800
A5: West      B5: 1200
A6: South     B6: 900
A7: East      B7: 1100
```

**Method 1 - Multiple SUMIF:**
```
=SUMIF(A:A, "West", B:B) + SUMIF(A:A, "East", B:B)
```
Returns: `4800` (West: 1000 + 1200 = 2200; East: 1500 + 1100 = 2600)

**Method 2 - SUMPRODUCT:**
```
=SUMPRODUCT(((A:A="West") + (A:A="East")), B:B)
```
Returns: `4800` (Same result as above)

**Method 3 - Array Formula:**
```
=SUM(IF((A:A="West") + (A:A="East"), B:B, 0))
```
Returns: `4800` (Same result, entered with Ctrl+Shift+Enter)

> [!NOTE]
> The SUMPRODUCT method is generally preferred as it doesn't require array entry and handles the logic more efficiently than multiple SUMIF calls for complex conditions.

> [!IMPORTANT]
> When using the `+` operator in SUMPRODUCT or array formulas for OR logic, be aware that if both conditions are TRUE for the same row, it will count that row only once, not twice.

> [!WARNING]
> For Method 3 (array formula with IF), remember to use Ctrl+Shift+Enter in Excel versions prior to 365. In Excel 365, this formula works as a regular formula due to dynamic arrays.

## Alternative Methods

### Using SUM with FILTER (Excel 365)
For modern Excel versions:
```
=SUM(FILTER(B:B, (A:A="West") + (A:A="East")))
```
This filters the data based on OR conditions before summing.

### Using SUMIFS with Array Constant
For simple text criteria:
```
=SUM(SUMIFS(B:B, A:A, {"West","East"}))
```
This passes an array of criteria to SUMIFS and sums the results.

### Using Boolean OR with Multiple Conditions
For more complex OR logic with different ranges:
```
=SUMPRODUCT(((A:A="West") + (C:C>100) + (D:D="Active")), B:B)
```
Sums column B where either column A is "West" OR column C > 100 OR column D is "Active".
