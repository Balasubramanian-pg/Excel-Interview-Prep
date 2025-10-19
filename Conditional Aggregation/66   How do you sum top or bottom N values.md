# How to Sum Top or Bottom N Values in Excel

This guide explains how to calculate the sum of the highest or lowest N values in a range using Excel functions, with methods compatible across different Excel versions.

## Formula Syntax

### Sum Top N Values
```
=SUMPRODUCT(LARGE(range, ROW(INDIRECT("1:"&N))))
```

### Sum Bottom N Values
```
=SUMPRODUCT(SMALL(range, ROW(INDIRECT("1:"&N))))
```

**Parameters:**
- `range`: The range containing numerical values to evaluate
- `N`: The number of top/bottom values to sum
- `ROW(INDIRECT("1:"&N))`: Creates an array of numbers from 1 to N

### Alternative Syntax (Excel 365)
```
=SUM(LARGE(range, SEQUENCE(N)))
```
```
=SUM(SMALL(range, SEQUENCE(N)))
```

## Worked Example

Given a range A1:A10 with values:
```
A1: 45
A2: 88
A3: 72
A4: 95
A5: 61
A6: 53
A7: 79
A8: 84
A9: 67
A10: 91
```

**Sum top 3 values:**
```
=SUMPRODUCT(LARGE(A1:A10, ROW(INDIRECT("1:3"))))
```
Returns: `274` (95 + 91 + 88)

**Sum bottom 3 values:**
```
=SUMPRODUCT(SMALL(A1:A10, ROW(INDIRECT("1:3"))))
```
Returns: `159` (45 + 53 + 61)

**Using explicit array (when N is small and known):**
```
=SUMPRODUCT(LARGE(A1:A10, {1;2;3}))
```
Returns: `274` (same result as above)

> [!NOTE]
> The ROW(INDIRECT("1:"&N)) construct creates a dynamic array of numbers from 1 to N, which the LARGE/SMALL functions use to return multiple values.

> [!IMPORTANT]
> When using the SUM(LARGE(range, ROW(1:N))) approach, you must enter it as an array formula (Ctrl+Shift+Enter) in older Excel versions. The SUMPRODUCT method doesn't require array entry.

> [!WARNING]
> If N exceeds the number of values in the range, these formulas will return #NUM! errors. Always ensure N is less than or equal to the count of numerical values in your range.

## Alternative Methods

### Using SUM with LARGE/SMALL (Array Formula)
For Excel versions requiring array formulas:
```
=SUM(LARGE(A1:A100, ROW(1:5)))
```
Enter with Ctrl+Shift+Enter to create an array formula.

### Using SUM with SEQUENCE (Excel 365)
For modern Excel versions with dynamic arrays:
```
=SUM(LARGE(A1:A100, SEQUENCE(5)))
```
This is the most straightforward approach in Excel 365.

### Using SUMIF with Threshold
For summing values above/below a certain threshold:
```
=SUMIF(A1:A100, ">"&LARGE(A1:A100, N+1))
```
This sums all values greater than the (N+1)th largest value.

---
