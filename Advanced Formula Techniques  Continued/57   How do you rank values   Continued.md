# Advanced Ranking in Excel

## Table of Contents
1. [Introduction](#introduction)
2. [RANK.AVG: Ranking with Averaged Ties](#rankavg-ranking-with-averaged-ties)
3. [PERCENTRANK.INC: Ranking as Percentile](#percentrankinc-ranking-as-percentile)
4. [Handling Duplicates: Unique Ranks](#handling-duplicates-unique-ranks)
5. [Step-by-Step Examples](#step-by-step-examples)
6. [Flashcard Q&A](#flashcard-qa)
7. [Best Practices and Tips](#best-practices-and-tips)
8. [Common Pitfalls and Warnings](#common-pitfalls-and-warnings)

---

## Introduction

- **Objective**: Assign ranks to values in a dataset, handling ties and percentiles.
- **Use Cases**: Competitive analysis, performance scoring, statistical reporting, and data normalization.
- **Functions Covered**:
  - `RANK.AVG`: Averages ranks for tied values.
  - `PERCENTRANK.INC`: Ranks values as percentiles.
  - Custom formula for unique ranks with duplicates.

---

## RANK.AVG: Ranking with Averaged Ties

### Syntax
```excel
=RANK.AVG(number, ref, [order])
```
- **`number`**: The value to rank.
- **`ref`**: The range of values.
- **`[order]`**: Optional. `0` or omitted for descending (default), `1` for ascending.

### Explanation
- **Ties**: If two or more values tie for the same rank, `RANK.AVG` assigns the **average** of their ranks.
- **Example**: If two values tie for 3rd place, both get `3.5`.

### Example
| A (Values) | B (Rank)          | B (Formula)                     |
|------------|-------------------|----------------------------------|
| 90         | 1                 | `=RANK.AVG(A1, $A$1:$A$5)`      |
| 85         | 2                 | `=RANK.AVG(A2, $A$1:$A$5)`      |
| 85         | 2                 | `=RANK.AVG(A3, $A$1:$A$5)`      |
| 80         | 4                 | `=RANK.AVG(A4, $A$1:$A$5)`      |
| 75         | 5                 | `=RANK.AVG(A5, $A$1:$A$5)`      |

> [!NOTE]
> `RANK.AVG` is useful for **fair ranking** when ties should share the same average rank.

---

## PERCENTRANK.INC: Ranking as Percentile

### Syntax
```excel
=PERCENTRANK.INC(array, x, [significance])
```
- **`array`**: The range of values.
- **`x`**: The value to rank.
- **`[significance]`**: Optional. Number of significant digits (default is 3).

### Explanation
- **Percentile Rank**: Returns the rank of a value as a **percentile** (0 to 1, inclusive).
- **Inclusive**: Includes the min and max values in the calculation.

### Example
| A (Values) | B (Percentile Rank) | B (Formula)                              |
|------------|---------------------|------------------------------------------|
| 90         | 1.000               | `=PERCENTRANK.INC($A$1:$A$5, A1, 3)`    |
| 85         | 0.600               | `=PERCENTRANK.INC($A$1:$A$5, A2, 3)`    |
| 85         | 0.600               | `=PERCENTRANK.INC($A$1:$A$5, A3, 3)`    |
| 80         | 0.400               | `=PERCENTRANK.INC($A$1:$A$5, A4, 3)`    |
| 75         | 0.200               | `=PERCENTRANK.INC($A$1:$A$5, A5, 3)`    |

> [!TIP]
> Use `PERCENTRANK.INC` for **normalizing data** or comparing values across different scales.

---

## Handling Duplicates: Unique Ranks

### Formula
```excel
=RANK.EQ(A1, $A$1:$A$100) + COUNTIF($A$1:A1, A1) - 1
```
- **`RANK.EQ`**: Assigns the same rank to ties.
- **`COUNTIF($A$1:A1, A1) - 1`**: Adjusts the rank to ensure uniqueness.

### Explanation
- **Unique Ranks**: Ensures each value gets a **unique rank**, even if there are duplicates.
- **How It Works**:
  - `RANK.EQ` gives the base rank.
  - `COUNTIF` counts how many times the current value has appeared so far.
  - Subtracting 1 adjusts the rank to avoid ties.

### Example
| A (Values) | B (Unique Rank)    | B (Formula)                                      |
|------------|--------------------|--------------------------------------------------|
| 90         | 1                  | `=RANK.EQ(A1, $A$1:$A$5) + COUNTIF($A$1:A1, A1) - 1` |
| 85         | 2                  | `=RANK.EQ(A2, $A$1:$A$5) + COUNTIF($A$1:A2, A2) - 1` |
| 85         | 3                  | `=RANK.EQ(A3, $A$1:$A$5) + COUNTIF($A$1:A3, A3) - 1` |
| 80         | 4                  | `=RANK.EQ(A4, $A$1:$A$5) + COUNTIF($A$1:A4, A4) - 1` |
| 75         | 5                  | `=RANK.EQ(A5, $A$1:$A$5) + COUNTIF($A$1:A5, A5) - 1` |

> [!IMPORTANT]
> This method ensures **no ties** in ranks, even with duplicate values.

---

## Step-by-Step Examples

### Example 1: RANK.AVG
1. **Data**: Column A contains `90`, `85`, `85`, `80`, `75`.
2. **Goal**: Rank values, averaging ties.
3. **Formula in B1**: `=RANK.AVG(A1, $A$1:$A$5)`
4. **Drag down**: B2: `=RANK.AVG(A2, $A$1:$A$5)`, etc.
5. **Result**: `1`, `2.5`, `2.5`, `4`, `5`

### Example 2: PERCENTRANK.INC
1. **Data**: Column A contains `90`, `85`, `85`, `80`, `75`.
2. **Goal**: Rank values as percentiles.
3. **Formula in B1**: `=PERCENTRANK.INC($A$1:$A$5, A1, 3)`
4. **Drag down**: B2: `=PERCENTRANK.INC($A$1:$A$5, A2, 3)`, etc.
5. **Result**: `1.000`, `0.600`, `0.600`, `0.400`, `0.200`

### Example 3: Unique Ranks
1. **Data**: Column A contains `90`, `85`, `85`, `80`, `75`.
2. **Goal**: Assign unique ranks, even for duplicates.
3. **Formula in B1**: `=RANK.EQ(A1, $A$1:$A$5) + COUNTIF($A$1:A1, A1) - 1`
4. **Drag down**: B2: `=RANK.EQ(A2, $A$1:$A$5) + COUNTIF($A$1:A2, A2) - 1`, etc.
5. **Result**: `1`, `2`, `3`, `4`, `5`

---

## Flashcard Q&A

### Q1: What does `RANK.AVG` do with tied values?
- **A**: Assigns the **average** of their ranks.

### Q2: How does `PERCENTRANK.INC` differ from `RANK.AVG`?
- **A**: `PERCENTRANK.INC` returns a **percentile** (0 to 1), while `RANK.AVG` returns a **rank**.

### Q3: How can you ensure unique ranks for duplicate values?
- **A**: Use `=RANK.EQ(A1, $A$1:$A$100) + COUNTIF($A$1:A1, A1) - 1`.

### Q4: What is the default order for `RANK.AVG`?
- **A**: **Descending** (highest value = rank 1).

---

## Best Practices and Tips

> [!TIP]
> - Use **`RANK.AVG`** for fair ranking with ties.
> - Use **`PERCENTRANK.INC`** for normalized comparisons.
> - For unique ranks, use the **custom formula** with `COUNTIF`.

> [!IMPORTANT]
> - Always **anchor ranges** with `$` (e.g., `$A$1:$A$100`).
> - Test ranking formulas on a **small dataset** first.

---

## Common Pitfalls and Warnings

> [!WARNING]
> - **`RANK.AVG` and `RANK.EQ`**: Can produce **different results** for ties.
> - **`PERCENTRANK.INC`**: Returns values between **0 and 1**, not ranks.

> [!CAUTION]
> - **Performance**: Using ranking formulas on large datasets can **slow down** calculations.

---

This document provides a **thorough, practical, and self-study guide** for advanced ranking in Excel, including handling ties, percentiles, and unique ranks.
