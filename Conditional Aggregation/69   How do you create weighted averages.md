# How to Create Weighted Averages in Excel

This guide explains how to calculate weighted averages in Excel, where values are multiplied by corresponding weights before averaging, giving more importance to some values than others.

## Formula Syntax

### Basic Weighted Average Formula
```
=SUMPRODUCT(values, weights) / SUM(weights)
```

**Parameters:**
- `values`: The range containing the values to be averaged
- `weights`: The range containing the corresponding weights for each value
- `SUMPRODUCT(values, weights)`: Calculates the sum of each value multiplied by its weight
- `SUM(weights)`: Calculates the total of all weights

## Worked Example

Given a dataset of student grades with different assignment weights:
```
A1: Assignment    B1: Score    C1: Weight
A2: Homework      B2: 85       C2: 0.2
A3: Quiz          B3: 92       C3: 0.3
A4: Midterm       B4: 78       C4: 0.25
A5: Final         B5: 88       C5: 0.25
```

**Calculate weighted average:**
```
=SUMPRODUCT(B2:B5, C2:C5) / SUM(C2:C5)
```

**Calculation breakdown:**
- SUMPRODUCT: (85×0.2) + (92×0.3) + (78×0.25) + (88×0.25) = 17 + 27.6 + 19.5 + 22 = 86.1
- SUM(weights): 0.2 + 0.3 + 0.25 + 0.25 = 1.0
- Result: 86.1 / 1.0 = 86.1

**Verification with regular average:**
```
=AVERAGE(B2:B5)
```
Returns: `85.75` (demonstrating how weights affect the result)

> [!NOTE]
> Weights don't need to sum to 1. The formula automatically normalizes by dividing by the total weight. If weights sum to 1, the division is essentially by 1.

> [!IMPORTANT]
> Ensure the values and weights ranges are the same size and properly aligned. SUMPRODUCT will return #VALUE! error if the ranges have different dimensions.

> [!TIP]
> Use named ranges to make your weighted average formulas more readable:
> `=SUMPRODUCT(Scores, Weights) / SUM(Weights)`

## Alternative Methods

### Using SUM with Array Multiplication
For compatibility with all Excel versions:
```
=SUM(values * weights) / SUM(weights)
```
Enter with Ctrl+Shift+Enter as an array formula in older Excel versions.

### Using SUMIF for Conditional Weighted Averages
To calculate weighted averages with conditions:
```
=SUMPRODUCT((criteria_range=criteria) * values, weights) / SUMIF(criteria_range, criteria, weights)
```
This calculates weighted average only for rows meeting specific criteria.

### Using AVERAGE.WEIGHTED (Excel 365)
In modern Excel versions with the new function:
```
=AVERAGE.WEIGHTED(values, weights)
```
This provides a more intuitive syntax for the same calculation.

### Handling Zero Weights
To avoid division by zero errors:
```
=IF(SUM(weights)>0, SUMPRODUCT(values, weights)/SUM(weights), 0)
```
This returns 0 if all weights are zero instead of a #DIV/0! error.
