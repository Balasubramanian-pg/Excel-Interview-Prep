# How to Sum with Multiple OR Criteria in Excel

This guide explains methods to sum values based on multiple OR conditions, where values are included if they meet at least one of several specified criteria.

## Formula Syntax

### Method 1: Multiple SUMIF Functions
```
=SUMIF(range, criteria1, sum_range) + SUMIF(range, criteria2, sum_range) + ...
```

**Parameters:**
- `range`: The range to evaluate for each condition
- `criteria1, criteria2, ...`: Individual conditions to check
- `sum_range`: The range containing values to sum
- Each SUMIF handles one condition, results are added together

### Method 2: SUMPRODUCT with Addition
```
=SUMPRODUCT(((range=criteria1) + (range=criteria2) + ...), sum_range)
```

**Parameters:**
- `range=criteria1`: Boolean array for first condition (returns TRUE/FALSE)
- `range=criteria2`: Boolean array for second condition
- `+` operator: Implements OR logic (TRUE if any condition is TRUE)
- `sum_range`: Values to sum

### Method 3: Array Formula with SUM
```
=SUM(IF((range=criteria1) + (range=criteria2), sum_range, 0))
```

**Parameters:**
- Requires Ctrl+Shift+Enter in older Excel versions
- `IF` function returns values where any condition is TRUE

## Worked Examples

Given sample sales data:
```
A1: Region   B1: Sales
A2: East     B2: 1000
A3: West     B3: 1500
A4: North    B4: 800
A5: South    B5: 1200
A6: East     B6: 900
A7: West     B7: 1100
A8: Central  B8: 1300
```

**Sum sales for East OR West regions (Multiple SUMIF):**
```
=SUMIF(A:A, "East", B:B) + SUMIF(A:A, "West", B:B)
```
Returns: `4500` (East: 1000 + 900 = 1900; West: 1500 + 1100 = 2600)

**Sum sales for East OR West regions (SUMPRODUCT):**
```
=SUMPRODUCT(((A:A="East") + (A:A="West")), B:B)
```
Returns: `4500` (Same result as above)

**Sum sales for East OR West OR Central regions:**
```
=SUMPRODUCT(((A:A="East") + (A:A="West") + (A:A="Central")), B:B)
```
Returns: `5800` (East: 1900 + West: 2600 + Central: 1300)

> [!NOTE]
> The SUMPRODUCT method is generally more efficient for multiple OR conditions as it processes the data in a single operation rather than multiple SUMIF calls.

> [!IMPORTANT]
> When using the `+` operator in SUMPRODUCT for OR logic, if multiple conditions are TRUE for the same row, that row's value will still be counted only once. The Boolean arrays convert to 1/0 values, so 1+1+0 = 2, but SUMPRODUCT handles this correctly.

> [!WARNING]
> For the array formula method (SUM with IF), remember to use Ctrl+Shift+Enter in Excel versions prior to 365. In Excel 365, this formula works as a regular formula due to dynamic arrays.

## Alternative Methods

### Using SUM with FILTER (Excel 365)
For modern Excel versions with dynamic arrays:
```
=SUM(FILTER(B:B, (A:A="East") + (A:A="West")))
```
Filters the data based on OR conditions before summing.

### Using SUMIFS with Array Constant
For simple text criteria with OR logic:
```
=SUM(SUMIFS(B:B, A:A, {"East","West"}))
```
Passes an array of criteria to SUMIFS and sums the results.

### Complex OR Logic with Different Ranges
For OR conditions across different columns:
```
=SUMPRODUCT(((A:A="East") + (C:C>100) + (D:D="Active")), B:B)
```
Sums column B where either:
- Column A is "East" OR
- Column C > 100 OR  
- Column D is "Active"

### Combining AND and OR Logic
For mixed logic conditions:
```
=SUMPRODUCT(((A:A="East") + (A:A="West")) * (C:C>50), B:B)
```
Sums column B where:
- (Region is "East" OR "West") AND
- Sales > 50

## Performance Considerations

**Use Multiple SUMIF when:**
- You have only 2-3 OR conditions
- Conditions are simple and easy to read separately
- Working with very large datasets where SUMPRODUCT might be slower

**Use SUMPRODUCT when:**
- You have multiple OR conditions (4+)
- Conditions are complex or involve calculations
- You want a single, compact formula
