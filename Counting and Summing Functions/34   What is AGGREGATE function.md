# AGGREGATE Function in Excel

This guide explains the AGGREGATE function, an advanced aggregation tool that combines the functionality of SUBTOTAL with enhanced error handling and additional mathematical operations.

## Function Syntax

### Reference Form (Ranges)
```
=AGGREGATE(function_num, options, ref1, [ref2], ...)
```

### Array Form (Arrays)
```
=AGGREGATE(function_num, options, array, [k])
```

**Parameters:**
- `function_num`: Number 1-19 specifying the aggregation function
- `options`: Number 0-7 specifying which values to ignore
- `ref1, ref2, ...`: Ranges to aggregate (reference form)
- `array`: Array or range to aggregate (array form)
- `k`: Required for functions that need a second parameter (LARGE, SMALL, etc.)

## Function Numbers

### Basic Aggregation (1-11)
```
1  = AVERAGE
2  = COUNT
3  = COUNTA
4  = MAX
5  = MIN
6  = PRODUCT
7  = STDEV.S
8  = STDEV.P
9  = SUM
10 = VAR.S
11 = VAR.P
```

### Advanced Functions (12-19)
```
12 = MEDIAN
13 = MODE.SNGL
14 = LARGE
15 = SMALL
16 = PERCENTILE.INC
17 = QUARTILE.INC
18 = PERCENTILE.EXC
19 = QUARTILE.EXC
```

## Options Parameter

The options parameter controls which values to ignore:
```
0 = Ignore nested SUBTOTAL and AGGREGATE functions
1 = Ignore hidden rows, nested SUBTOTAL and AGGREGATE functions
2 = Ignore error values, nested SUBTOTAL and AGGREGATE functions
3 = Ignore hidden rows, error values, nested SUBTOTAL and AGGREGATE functions
4 = Ignore nothing
5 = Ignore hidden rows
6 = Ignore error values
7 = Ignore hidden rows and error values
```

## Worked Examples

Given sample data with some issues:
```
A1: Sales
A2: 1000
A3: #DIV/0!
A4: 1500
A5: 
A6: 2000
A7: #N/A
A8: 1200
```

**Basic sum ignoring errors:**
```
=AGGREGATE(9, 6, A1:A8)
```
Returns: `5700` (sum of 1000 + 1500 + 2000 + 1200, ignoring errors and empty cells)

**Average ignoring errors and hidden rows:**
```
=AGGREGATE(1, 7, A1:A8)
```
Returns: `1425` (average of visible, non-error values)

**Maximum value ignoring errors:**
```
=AGGREGATE(4, 6, A1:A8)
```
Returns: `2000` (maximum value ignoring errors)

**Second largest value:**
```
=AGGREGATE(14, 6, A1:A8, 2)
```
Returns: `1500` (second largest value ignoring errors, function 14 = LARGE)

**Percentile calculation:**
```
=AGGREGATE(16, 6, A1:A8, 0.25)
```
Returns: `1100` (25th percentile of values ignoring errors)

> [!NOTE]
> AGGREGATE was introduced in Excel 2010 and provides significant advantages over SUBTOTAL, particularly in error handling and additional mathematical functions.

> [!IMPORTANT]
> When using functions that require the [k] parameter (LARGE, SMALL, PERCENTILE, QUARTILE), you must use the array form syntax. The reference form only works for functions 1-13.

> [!TIP]
> Use option 6 or 7 to handle datasets with potential errors gracefully. This prevents entire formulas from failing when only some cells contain errors.

## Comparison with SUBTOTAL

### AGGREGATE Advantages
- More function options (19 vs 11)
- Better error handling
- Can ignore nested SUBTOTAL/AGGREGATE functions
- Includes statistical functions like MEDIAN, PERCENTILE

### SUBTOTAL Advantages
- Simpler syntax for basic needs
- Better compatibility with older Excel versions
- Automatically updates when filters change

## Practical Applications

### Handling Dirty Data
```
=AGGREGATE(1, 7, SalesData)
```
Calculates average while ignoring errors and hidden rows in potentially problematic data.

### Robust Large/Small Calculations
```
=AGGREGATE(14, 6, A:A, 3)
```
Finds the third largest value in column A, ignoring any errors that would cause LARGE to fail.

### Conditional Aggregation
Combine with other functions for conditional analysis:
```
=AGGREGATE(9, 6, (A1:A10)*(B1:B10="Criteria"))
```
Sums values in A1:A10 where corresponding cells in B1:B10 meet criteria, ignoring errors.
