# How to Count Unique Values with Criteria in Excel

This guide explains methods to count distinct values that meet specific conditions, with solutions for both modern Excel versions and older versions requiring array formulas.

## Formula Syntax

### Excel 365 Method (Dynamic Arrays)
```
=COUNTA(UNIQUE(FILTER(data_range, criteria_range="criteria")))
```

**Parameters:**
- `data_range`: The range containing values to check for uniqueness
- `criteria_range`: The range to evaluate against the criteria
- `"criteria"`: The condition that must be met
- `FILTER`: Returns only the values that meet the criteria
- `UNIQUE`: Extracts distinct values from the filtered results
- `COUNTA`: Counts the number of unique values

### Older Excel Method (Array Formula)
```
=SUM(IF(criteria_range="criteria", 1/COUNTIFS(data_range, data_range, criteria_range, "criteria"), 0))
```

**Parameters:**
- Requires Ctrl+Shift+Enter in pre-365 versions
- `1/COUNTIFS(...)`: Creates fractions that sum to 1 for each unique value
- `SUM(IF(...))`: Sums the fractions to count unique values

## Worked Example

Given a dataset of sales transactions:
```
A1: Product    B1: Region
A2: Apple      B2: North
A3: Orange     B3: South
A4: Apple      B4: North
A5: Banana     B5: North
A6: Orange     B6: North
A7: Apple      B7: South
A8: Banana     B8: South
```

**Count unique products in North region (Excel 365):**
```
=COUNTA(UNIQUE(FILTER(A:A, B:B="North")))
```

**Calculation breakdown:**
- FILTER returns: {"Apple"; "Orange"; "Apple"; "Banana"; "Orange"}
- UNIQUE returns: {"Apple"; "Orange"; "Banana"}
- COUNTA returns: `3`

**Count unique products in North region (Older Excel):**
```
=SUM(IF(B2:B8="North", 1/COUNTIFS(A2:A8, A2:A8, B2:B8, "North"), 0))
```
Returns: `3` (entered with Ctrl+Shift+Enter)

> [!NOTE]
> The Excel 365 method is more intuitive and easier to debug since you can evaluate each function step by step. The older method uses a mathematical approach with frequency distributions.

> [!IMPORTANT]
> In the older Excel method, the formula must be entered as an array formula (Ctrl+Shift+Enter) in versions prior to Excel 365. Excel 365 handles this as a regular formula due to dynamic arrays.

> [!WARNING]
> Both methods count blank cells as unique values. To exclude blanks, add an additional criteria: `FILTER(A:A, (B:B="North")*(A:A<>""))` or modify the array formula to exclude empty cells.

## Alternative Methods

### Using SUMPRODUCT (Older Excel Alternative)
For a non-array formula approach in older versions:
```
=SUMPRODUCT((criteria_range="criteria")/COUNTIFS(data_range, data_range, criteria_range, criteria_range))
```
This uses SUMPRODUCT to avoid array formula entry.

### Multiple Criteria with Excel 365
To count unique values with multiple conditions:
```
=COUNTA(UNIQUE(FILTER(data_range, (criteria_range1="criteria1")*(criteria_range2="criteria2"))))
```

### Using PivotTables
For visual analysis:
- Create a PivotTable with the data range as rows
- Add the criteria field as a filter
- The row count in the PivotTable shows unique values meeting the filter criteria

### Handling Text and Numbers Separately
If your data contains mixed types and you want to count only text values:
```
=COUNTA(UNIQUE(FILTER(data_range, (criteria_range="criteria")*(ISTEXT(data_range)))))
```
