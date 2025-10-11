# Creating Running Totals in Excel

## Table of Contents
1. [Introduction](#introduction)
2. [Method 1: Simple Running Total](#method-1-simple-running-total)
3. [Method 2: Running Total with SUMIF for Grouped Data](#method-2-running-total-with-sumif-for-grouped-data)
4. [Method 3: Running Total with SCAN (Excel 365)](#method-3-running-total-with-scan-excel-365)
5. [Step-by-Step Examples](#step-by-step-examples)
6. [Flashcard Q&A](#flashcard-qa)
7. [Best Practices and Tips](#best-practices-and-tips)
8. [Common Pitfalls and Warnings](#common-pitfalls-and-warnings)

---

## Introduction

- **Objective**: Calculate a cumulative sum (running total) of values in a column or range.
- **Use Cases**: Financial reports, inventory tracking, progress monitoring, and data analysis.
- **Methods**:
  - Simple expanding range (`SUM`).
  - Grouped data (`SUMIF`).
  - Dynamic arrays (`SCAN` in Excel 365).

---

## Method 1: Simple Running Total

### Syntax
```excel
=SUM($A$1:A1)
```
- **`$A$1:A1`**: Expanding range. `$A$1` is fixed, `A1` changes as you drag the formula down.

### Explanation
- **`$A$1`**: Absolute reference to the first cell.
- **`A1`**: Relative reference that expands as you drag the formula down.
- **Result**: Each row shows the sum of all values from the first row to the current row.

### Example
| A (Values) | B (Running Total) | B (Formula)         |
|------------|-------------------|---------------------|
| 10         | 10                | `=SUM($A$1:A1)`     |
| 20         | 30                | `=SUM($A$1:A2)`     |
| 30         | 60                | `=SUM($A$1:A3)`     |

> [!NOTE]
> This method is **simple and widely compatible** with all Excel versions.

---

## Method 2: Running Total with SUMIF for Grouped Data

### Syntax
```excel
=SUMIF($A$1:A1, A1, $B$1:B1)
```
- **`$A$1:A1`**: Expanding range for criteria.
- **`A1`**: Current row’s group identifier.
- **`$B$1:B1`**: Expanding range for values to sum.

### Explanation
- **`SUMIF`**: Sums values in `$B$1:B1` where the corresponding cell in `$A$1:A1` matches `A1`.
- **Use Case**: Calculate running totals for specific groups or categories.

### Example
| A (Group) | B (Value) | C (Running Total) | C (Formula)                     |
|-----------|-----------|-------------------|---------------------------------|
| Fruit     | 10        | 10                | `=SUMIF($A$1:A1, A1, $B$1:B1)`  |
| Fruit     | 20        | 30                | `=SUMIF($A$1:A2, A2, $B$1:B2)`  |
| Veg       | 5         | 5                 | `=SUMIF($A$1:A3, A3, $B$1:B3)`  |
| Fruit     | 30        | 60                | `=SUMIF($A$1:A4, A4, $B$1:B4)`  |

> [!TIP]
> Useful for **grouped or categorical data**.

---

## Method 3: Running Total with SCAN (Excel 365)

### Syntax
```excel
=SCAN(0, A1:A100, LAMBDA(acc, val, acc + val))
```
- **`0`**: Initial value for the accumulator.
- **`A1:A100`**: Range of values.
- **`LAMBDA(acc, val, acc + val)`**: Lambda function to add each value to the accumulator.

### Explanation
- **`SCAN`**: Applies the lambda function to each value in the range, returning an array of running totals.
- **Dynamic Array**: Spills results automatically in Excel 365.

### Example
| A (Values) | B (Running Total) | B (Formula)                                      |
|------------|-------------------|--------------------------------------------------|
| 10         | 10                | `=SCAN(0, A1:A3, LAMBDA(acc, val, acc + val))`    |
| 20         | 30                |                                                  |
| 30         | 60                |                                                  |

> [!IMPORTANT]
> Requires **Excel 365** or later. Spills results dynamically.

---

## Step-by-Step Examples

### Example 1: Simple Running Total
1. **Data**: Column A contains `10`, `20`, `30`.
2. **Goal**: Calculate running total in column B.
3. **Formula in B1**: `=SUM($A$1:A1)`
4. **Drag down**: B2: `=SUM($A$1:A2)`, B3: `=SUM($A$1:A3)`
5. **Result**: `10`, `30`, `60`

### Example 2: Grouped Running Total
1. **Data**: Column A contains `Fruit`, `Fruit`, `Veg`, `Fruit`; Column B contains `10`, `20`, `5`, `30`.
2. **Goal**: Calculate running total for each group in column C.
3. **Formula in C1**: `=SUMIF($A$1:A1, A1, $B$1:B1)`
4. **Drag down**: C2: `=SUMIF($A$1:A2, A2, $B$1:B2)`, etc.
5. **Result**: `10`, `30`, `5`, `60`

### Example 3: SCAN Running Total
1. **Data**: Column A contains `10`, `20`, `30`.
2. **Goal**: Calculate running total in column B.
3. **Formula in B1**: `=SCAN(0, A1:A3, LAMBDA(acc, val, acc + val))`
4. **Result**: Spills `10`, `30`, `60` in B1:B3.

---

## Flashcard Q&A

### Q1: What is a running total?
- **A**: A cumulative sum of values up to the current row.

### Q2: How does the simple running total formula work?
- **A**: It uses an expanding range (`$A$1:A1`) to sum all values from the first row to the current row.

### Q3: When should you use `SUMIF` for running totals?
- **A**: When you need to calculate running totals for **grouped or categorical data**.

### Q4: What is the advantage of using `SCAN` for running totals?
- **A**: It dynamically spills results and is **easier to maintain** in Excel 365.

---

## Best Practices and Tips

> [!TIP]
> - Use **named ranges** for clarity.
> - For large datasets, **avoid volatile functions** like `OFFSET`.
> - In Excel 365, **`SCAN`** is the most efficient and flexible method.

> [!IMPORTANT]
> - Test formulas on a **small dataset** first.
> - Use **absolute references** (`$A$1`) to anchor the starting cell.

---

## Common Pitfalls and Warnings

> [!WARNING]
> - **Simple Method**: Dragging the formula incorrectly can lead to **wrong ranges**.
> - **SUMIF Method**: Ensure the criteria range and sum range are **aligned**.

> [!CAUTION]
> - **Performance**: Using entire columns (e.g., `A:A`) in large workbooks can **slow down** calculations.

---

This document provides a **thorough, practical, and self-study guide** for creating running totals in Excel, covering all methods, examples, and best practices.
