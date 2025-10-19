# How to Sum with Multiple AND Conditions in Excel

This guide explains how to sum values based on multiple criteria that must all be true using Excel's SUMIFS function, which is specifically designed for AND logic across multiple conditions.

## Formula Syntax

### SUMIFS Function
```
=SUMIFS(sum_range, criteria_range1, criteria1, criteria_range2, criteria2, ...)
```

**Parameters:**
- `sum_range`: The range containing numerical values to sum
- `criteria_range1`: The first range to evaluate against criteria1
- `criteria1`: The condition that must be met in criteria_range1
- `criteria_range2`, `criteria2`: Additional criteria ranges and conditions (up to 127 pairs)

## Worked Example

Given the following dataset:
```
A1: Region    B1: Sales    C1: Status    D1: Revenue
A2: West      B2: 1500     C2: Active    D2: 25000
A3: East      B3: 800      C3: Active    D3: 12000
A4: West      B4: 2000     C4: Inactive  D4: 35000
A5: North     B5: 1200     C5: Active    D5: 18000
A6: West      B6: 1800     C6: Active    D6: 30000
```

**Sum with multiple AND conditions:**
```
=SUMIFS(D:D, A:A, "West", B:B, ">1000", C:C, "Active")
```

Returns: `55000` (Only D2 and D6 meet all three conditions: Region = "West", Sales > 1000, and Status = "Active")

**Breakdown of matching rows:**
- Row 2: West, 1500, Active → MATCH (25000)
- Row 4: West, 2000, Inactive → NO MATCH (Status ≠ "Active")
- Row 6: West, 1800, Active → MATCH (30000)

> [!NOTE]
> SUMIFS was introduced in Excel 2007 and is the preferred method for multiple AND conditions. The older SUMPRODUCT method still works but is less efficient.

> [!IMPORTANT]
> The order of arguments in SUMIFS is different from SUMIF:
> - SUMIF: `=SUMIF(criteria_range, criteria, sum_range)`
> - SUMIFS: `=SUMIFS(sum_range, criteria_range1, criteria1, ...)`

> [!TIP]
> Use cell references for criteria to make formulas dynamic:
> `=SUMIFS(D:D, A:A, F1, B:B, ">"&F2, C:C, F3)`
> Where F1, F2, F3 contain the criteria values.

## Alternative Methods

### Using SUMPRODUCT (Pre-2007 Excel)
For compatibility with older Excel versions:
```
=SUMPRODUCT((A:A="West")*(B:B>1000)*(C:C="Active")*(D:D))
```
This uses Boolean logic to achieve the same result as SUMIFS.

### Using SUM with FILTER (Excel 365)
For modern Excel versions with dynamic arrays:
```
=SUM(FILTER(D:D, (A:A="West")*(B:B>1000)*(C:C="Active")))
```
This filters the data first, then sums the results.

### Using Database Functions
For complex criteria sets:
```
=DSUM(A1:D100, "Revenue", F1:H2)
```
Where F1:H2 contains a criteria range with field headers and conditions.
