# How to Count Cells with Specific Text in Excel

This guide explains how to count cells containing specific text in Excel using various matching criteria, including exact matches, partial matches, and case-sensitive counts.

## Formula Syntax

### COUNTIF for Text Matching
```
=COUNTIF(range, criteria)
```

**Parameters:**
- `range`: The range of cells to evaluate
- `criteria`: The text pattern to count, which can include wildcards (*) for partial matching

### SUMPRODUCT with EXACT for Case-Sensitive Counting
```
=SUMPRODUCT(--EXACT(range, "text"))
```

**Parameters:**
- `range`: The range of cells to evaluate
- `"text"`: The exact text to match (case-sensitive)

## Worked Example

Given a column A with the following data:
```
A1: Apple
A2: apple
A3: Pineapple
A4: APPLE
A5: Orange
A6: Green Apple
```

**Exact match:**
```
=COUNTIF(A:A, "Apple")
```
Returns: `1` (only A1 matches exactly "Apple")

**Contains text (wildcard):**
```
=COUNTIF(A:A, "*apple*")
```
Returns: `4` (A1, A2, A3, A6 contain "apple")

**Starts with:**
```
=COUNTIF(A:A, "apple*")
```
Returns: `1` (only A2 starts with "apple")

**Ends with:**
```
=COUNTIF(A:A, "*apple")
```
Returns: `2` (A1 and A2 end with "apple")

**Case-sensitive count:**
```
=SUMPRODUCT(--EXACT(A1:A100, "Apple"))
```
Returns: `1` (only A1 matches "Apple" exactly with same case)

> [!NOTE]
> The COUNTIF function is not case-sensitive. For case-sensitive counting, you must use the SUMPRODUCT/EXACT combination or other array formulas.

> [!IMPORTANT]
> Wildcards in COUNTIF:
> - `*` represents any number of characters
> - `?` represents a single character
> - Use `~` to escape wildcards if you need to search for literal * or ? characters

> [!WARNING]
> When using SUMPRODUCT with EXACT, ensure you reference the same sized ranges. The formula will return errors if the ranges are not compatible.

## Alternative Methods

### Using COUNTIFS for Multiple Criteria
For counting cells that meet multiple text conditions:
```
=COUNTIFS(A:A, "*apple*", B:B, ">100")
```
This counts cells in column A containing "apple" where corresponding cells in column B are greater than 100.

### Using FILTER function (Excel 365)
For modern Excel versions, you can use FILTER to count text matches:
```
=COUNTA(FILTER(A:A, ISNUMBER(SEARCH("apple", A:A))))
```
This provides an alternative method for partial text matching.
