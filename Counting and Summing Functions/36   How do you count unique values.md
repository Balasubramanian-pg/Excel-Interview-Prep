# How to Count Unique Values in Excel

This guide explains multiple methods to count distinct values in a range, with solutions for both modern Excel versions with dynamic arrays and older versions requiring array formulas.

## Formula Syntax

### Excel 365 Method (Dynamic Arrays)
```
=COUNTA(UNIQUE(range))
```

**Parameters:**
- `range`: The range containing values to evaluate for uniqueness
- `UNIQUE`: Extracts distinct values from the range
- `COUNTA`: Counts the number of unique values returned

### Older Excel Method (Frequency Distribution)
```
=SUMPRODUCT(1/COUNTIF(range, range))
```

**Parameters:**
- `range`: The range containing values to evaluate
- `COUNTIF(range, range)`: Creates an array counting occurrences of each value
- `1/COUNTIF(...)`: Creates fractions that sum to 1 for each unique value
- `SUMPRODUCT`: Sums the fractions to count unique values

### Array Formula Alternative
```
=SUM(1/COUNTIF(range, range))
```

**Parameters:**
- Same logic as SUMPRODUCT method
- Requires Ctrl+Shift+Enter in older Excel versions

## Worked Examples

Given sample data in range A1:A10:
```
A1: Apple
A2: Orange
A3: Apple
A4: Banana
A5: Orange
A6: Apple
A7: Peach
A8: Banana
A9: Orange
A10: Apple
```

**Excel 365 Method:**
```
=COUNTA(UNIQUE(A1:A10))
```

**Calculation breakdown:**
- UNIQUE returns: {"Apple"; "Orange"; "Banana"; "Peach"}
- COUNTA returns: `4`

**Older Excel Method:**
```
=SUMPRODUCT(1/COUNTIF(A1:A10, A1:A10))
```

**Calculation breakdown:**
- COUNTIF for each position: {4;3;4;2;3;4;1;2;3;4} (occurrences of each value)
- 1/COUNTIF: {0.25;0.333;0.25;0.5;0.333;0.25;1;0.5;0.333;0.25}
- SUMPRODUCT sum: `0.25+0.333+0.25+0.5+0.333+0.25+1+0.5+0.333+0.25 = 4`

**Array Formula Method:**
```
=SUM(1/COUNTIF(A1:A10, A1:A10))
```
Returns: `4` (entered with Ctrl+Shift+Enter in older Excel)

> [!NOTE]
> The Excel 365 method is more intuitive and efficient. The older method uses a mathematical approach where each occurrence of a value contributes a fraction (1/n), so n occurrences of the same value sum to 1.

> [!IMPORTANT]
> The frequency distribution method (SUMPRODUCT/COUNTIF) will return #DIV/0! if the range contains blank cells. To handle blanks, use: `=SUMPRODUCT((range<>"")/COUNTIF(range, range&""))`

> [!WARNING]
> For the SUM array formula method, remember to use Ctrl+Shift+Enter in Excel versions prior to 365. Excel 365 handles this as a regular formula due to dynamic arrays.

## Handling Special Cases

### Excluding Blank Cells
**Excel 365:**
```
=COUNTA(UNIQUE(FILTER(range, range<>"")))
```

**Older Excel:**
```
=SUMPRODUCT((range<>"")/COUNTIF(range, range&""))
```

### Counting Unique Numeric Values Only
**Excel 365:**
```
=COUNT(UNIQUE(FILTER(range, ISNUMBER(range))))
```

**Older Excel:**
```
=SUMPRODUCT((ISNUMBER(range))/COUNTIF(range, range))
```

### Case-Sensitive Unique Count
```
=SUMPRODUCT(1/COUNTIFS(range, range, range, "<>"))
```
Note: COUNTIFS with additional criteria can help with more complex scenarios

## Alternative Methods

### Using FREQUENCY for Numeric Values
For counting unique numbers only:
```
=SUM(IF(FREQUENCY(range, range)>0, 1))
```
Array formula that works well with numeric data

### Using PivotTables
For visual analysis:
- Create a PivotTable with the data range as Rows
- The row count in the PivotTable shows unique values
- Use "Distinct Count" in Values area if available

### Using Advanced Filter
Manual method:
1. Select Data > Advanced Filter
2. Choose "Unique records only"
3. Count the filtered results

### Excel 365 Dynamic Array Alternatives
```
=ROWS(UNIQUE(range))
```
Alternative to COUNTA(UNIQUE()) when you only need the count

```
=COUNT(UNIQUE(range))
```
Counts unique numeric values specifically

