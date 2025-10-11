# Removing Duplicates with Formulas in Excel

## Table of Contents
1. [Introduction](#introduction)
2. [Excel 365: Using UNIQUE](#excel-365-using-unique)
3. [Older Excel: Array Formula](#older-excel-array-formula)
4. [Step-by-Step Examples](#step-by-step-examples)
5. [Flashcard Q&A](#flashcard-qa)
6. [Best Practices and Tips](#best-practices-and-tips)
7. [Common Pitfalls and Warnings](#common-pitfalls-and-warnings)

---

## Introduction

- **Purpose**: Remove duplicate values from a list using formulas.
- **Why Use It?**: Maintain data integrity, simplify analysis, and automate reporting.
- **Compatibility**:
  - **Excel 365**: Uses the `UNIQUE` function.
  - **Older Excel**: Uses an array formula with `INDEX`, `MATCH`, and `COUNTIF`.

---

## Excel 365: Using UNIQUE

### Syntax
```excel
=UNIQUE(array)
```
- **`array`**: The range or array from which to remove duplicates.

### How It Works
- **Behavior**: Returns a **spilled array** of unique values.
- **Example**:
  ```excel
  =UNIQUE(A1:A100)
  ```
  - If `A1:A100` contains duplicates, `UNIQUE` returns only the unique values.

### Example
| A (Data)       | B (Formula)         | B (Result)          |
|----------------|----------------------|---------------------|
| Apple          | `=UNIQUE(A1:A5)`     | Apple               |
| Banana         |                      | Banana              |
| Apple          |                      | Orange              |
| Orange         |                      | Grape               |
| Grape          |                      |                     |

- **Result**: `B1:B4` contains `Apple`, `Banana`, `Orange`, `Grape`.

> [!NOTE]
> `UNIQUE` is **dynamic** and automatically updates when data changes.

---

## Older Excel: Array Formula

### Syntax
```excel
=INDEX($A$1:$A$100, MATCH(0, COUNTIF($B$1:B1, $A$1:$A$100), 0))
```
- **`$A$1:$A$100`**: The range to check for duplicates.
- **`$B$1:B1`**: The range where unique values are listed (expands as you drag down).

### How It Works
- **`COUNTIF($B$1:B1, $A$1:$A$100)`**: Counts how many times each value in `A1:A100` appears in `B1:B1`.
- **`MATCH(0, ..., 0)`**: Finds the first value in `A1:A100` that hasn’t been listed in `B1:B1`.
- **`INDEX`**: Returns the value at that position.

### Example
| A (Data)       | B (Formula)                                      | B (Result)          |
|----------------|---------------------------------------------------|---------------------|
| Apple          | `=INDEX($A$1:$A$5, MATCH(0, COUNTIF($B$1:B1, $A$1:$A$5), 0))` | Apple               |
| Banana         | Drag down                                         | Banana              |
| Apple          |                                                   | Orange              |
| Orange         |                                                   | Grape               |
| Grape          |                                                   |                     |

- **Result**: `B1:B4` contains `Apple`, `Banana`, `Orange`, `Grape`.

> [!IMPORTANT]
> - Enter the formula in `B1` and **drag down**.
> - Use **Ctrl+Shift+Enter** if not in Excel 365.

---

## Step-by-Step Examples

### Example 1: Excel 365 (UNIQUE)
1. **Data**: Column A contains `Apple`, `Banana`, `Apple`, `Orange`, `Grape`.
2. **Goal**: List unique values in column B.
3. **Formula in B1**: `=UNIQUE(A1:A5)`
4. **Result**: `B1:B4` spills `Apple`, `Banana`, `Orange`, `Grape`.

### Example 2: Older Excel (Array Formula)
1. **Data**: Column A contains `Apple`, `Banana`, `Apple`, `Orange`, `Grape`.
2. **Goal**: List unique values in column B.
3. **Formula in B1**:
   ```excel
   =INDEX($A$1:$A$5, MATCH(0, COUNTIF($B$1:B1, $A$1:$A$5), 0))
   ```
4. **Drag down**: `B2`, `B3`, etc.
5. **Result**: `B1:B4` contains `Apple`, `Banana`, `Orange`, `Grape`.

---

## Flashcard Q&A

### Q1: How do you remove duplicates in Excel 365?
- **A**: Use `=UNIQUE(A1:A100)`.

### Q2: How does the array formula for removing duplicates work in older Excel?
- **A**: It uses `INDEX`, `MATCH`, and `COUNTIF` to find the next unique value.

### Q3: What is the advantage of `UNIQUE` in Excel 365?
- **A**: It **automatically spills** and updates with data changes.

### Q4: Why is `COUNTIF($B$1:B1, $A$1:$A$100)` used in the array formula?
- **A**: It counts how many times each value in `A1:A100` has already been listed in `B1:B1`.

---

## Best Practices and Tips

> [!TIP]
> - Use `UNIQUE` in **Excel 365** for simplicity and dynamic updates.
> - In older Excel, **drag the formula down** to list all unique values.
> - Combine with **`SORT`** in Excel 365 to sort unique values:
>   ```excel
   =SORT(UNIQUE(A1:A100))
   ```

> [!IMPORTANT]
> - Test formulas on a **small dataset** first.
> - Use **absolute references** (e.g., `$A$1:$A$100`) in the array formula.

---

## Common Pitfalls and Warnings

> [!WARNING]
> - **Legacy Arrays**: Forgetting to **drag down** the formula in older Excel.
> - **Spill Errors**: In Excel 365, ensure the spill range is **clear** to avoid `#SPILL!` errors.

> [!CAUTION]
> - **Performance**: Using array formulas on large datasets can **slow down** calculations.
> - **Compatibility**: `UNIQUE` is **only available in Excel 365**.

---

This document provides a **detailed, practical, and self-study guide** for removing duplicates using formulas in Excel, covering both modern and legacy methods.
