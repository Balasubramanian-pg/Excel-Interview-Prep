### 137. How do you create dynamic arrays that resize?

Dynamic arrays, a feature in Excel 365 and Excel 2021, allow a single formula to return (or "spill") results into multiple cells. The key benefit is that the size of this spill range automatically adjusts as your source data changes. This eliminates the need for old array formulas (`Ctrl+Shift+Enter`) or manually dragging formulas down.

#### Basic Resizing Array with `FILTER`

The simplest way to create a dynamic copy of your data that resizes is by using the `FILTER` function. The formula filters your entire dataset based on a condition that includes all non-empty rows.

**Formula:**
```excel
=FILTER(A:C, A:A<>"")
```

**How it works:**
*   `FILTER(A:C, ...)`: This tells Excel you want to return a subset of the data from columns A, B, and C.
*   `... A:A<>""`: This is the condition. It checks every cell in column A and includes the row in the result if the cell is not empty (`<>""`).

This single formula, entered in one cell, will spill to create a complete copy of your data. When you add a new row to the source data (in columns A:C), the spilled array will automatically expand to include it.

> [!NOTE]
> For this formula to work reliably, your data should not have blank rows interspersed within it. Using an Excel Table for your source data is the recommended best practice.

#### Adding Dynamic Row Numbers

You can enhance your dynamic array by adding a column of sequential row numbers that also updates automatically. This is done by combining `SEQUENCE`, `COUNTA`, and `HSTACK`.

**Formula:**
```excel
=HSTACK(SEQUENCE(COUNTA(A:A)), FILTER(A:C, A:A<>""))
```

**How it works:**
1.  `COUNTA(A:A)`: First, this function counts the number of non-empty cells in column A to determine how many rows of data you have.
2.  `SEQUENCE(COUNTA(A:A))`: The `SEQUENCE` function then generates a dynamic vertical list of numbers, starting from 1 up to the total count from the previous step.
3.  `FILTER(A:C, A:A<>"")`: This is the same resizing array from the first example.
4.  `HSTACK(...)`: The `HSTACK` function takes the two arrays (the sequence of numbers and the filtered data) and stacks them side-by-side ("horizontally") into a single new array.

> [!TIP]
> **Use Excel Tables for Robustness**
> Converting your source data range to an official Excel Table (select your data and press `Ctrl+T`) makes these formulas even more powerful and readable. Tables automatically manage their own size.
>
> For example, if your table is named `SalesData`, the formula becomes:
> `=FILTER(SalesData, SalesData[Region]<>"")`
> This formula is easier to understand and less prone to error than using whole-column references like `A:A`.
