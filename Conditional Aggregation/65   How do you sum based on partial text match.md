# How to Sum Based on Partial Text Match in Excel

This guide explains how to sum values based on partial text matches using wildcards in Excel's SUMIF function, including methods for handling multiple partial match conditions.

## Formula Syntax

### Basic SUMIF with Wildcard
```
=SUMIF(range, criteria, sum_range)
```

**Parameters:**
- `range`: The range to check for text matches
- `criteria`: The text pattern to match, using wildcards for partial matching
- `sum_range`: The range containing values to sum

### Multiple Partial Matches (Additive Approach)
```
=SUMIF(range, "*text1*", sum_range) + SUMIF(range, "*text2*", sum_range)
```

**Parameters:**
- Each SUMIF function handles a different text pattern
- Results are added together for a cumulative total

## Worked Example

Given the following data:
```
A1: West Region    B1: 100
A2: East Region    B2: 150
A3: North West     B3: 200
A4: South East     B4: 75
A5: Central        B5: 125
```

**Single partial match:**
```
=SUMIF(A:A, "*West*", B:B)
```
Returns: `300` (B1 + B3 = 100 + 200)

**Multiple partial matches:**
```
=SUMIF(A:A, "*West*", B:B) + SUMIF(A:A, "*East*", B:B)
```
Returns: `525` (West matches: 300 + East matches: 225)

> [!NOTE]
> The asterisk (*) wildcard represents any sequence of characters. Use "*West*" to match any text containing "West" anywhere in the string.

> [!IMPORTANT]
> SUMIF requires three arguments when summing a different range than the criteria range. If omitted, Excel will sum the criteria range instead of the intended sum range.

> [!TIP]
- Use "West*" to match text starting with "West"
- Use "*West" to match text ending with "West"  
- Use "??West" to match text where "West" is preceded by exactly two characters

## Alternative Methods

### Using SUMIFS for Multiple Criteria
For more complex conditions with multiple criteria:
```
=SUMIFS(B:B, A:A, "*West*", C:C, ">50")
```
This sums column B where column A contains "West" AND column C is greater than 50.

### Using SUMPRODUCT with SEARCH (Case-Sensitive Alternative)
```
=SUMPRODUCT(B1:B100, --(ISNUMBER(SEARCH("West", A1:A100))))
```
This provides a case-sensitive alternative to SUMIF with wildcards.

### Using FILTER function (Excel 365)
For modern Excel versions:
```
=SUM(FILTER(B:B, ISNUMBER(SEARCH("West", A:A))))
```
This uses FILTER to extract matching values before summing.

---
