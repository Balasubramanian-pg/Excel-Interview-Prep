### 120. How do you create running differences (deltas)?

Calculating running differences, or "deltas," is essential for analyzing trends, momentum, and the rate of change in a dataset over time. These formulas compare a data point to the value from a previous period.

#### Simple Difference (Period-over-Period)

This is the most basic delta calculation, showing the absolute change between consecutive values.

**Formula:**
```excel
=A2-A1
```

**How it works:**
This formula subtracts the previous period's value (`A1`) from the current period's value (`A2`). The result is the raw increase or decrease.

> [!IMPORTANT]
> This formula must be entered in the **second row** of your data, as the first row has no preceding value to compare against. You then drag it down the column.

#### Percentage Change

This formula provides more context by showing the change as a percentage of the previous period's value, which is useful for comparing growth rates across different scales.

**Formula:**
```excel
=(A2-A1)/A1
```

**How it works:**
It first calculates the simple difference (`A2-A1`) and then divides it by the original value (`A1`) to get the relative change. The cell must be formatted as a **Percentage**.

> [!CAUTION]
> This formula will return a `#DIV/0!` error if the previous period's value (`A1`) is zero. To prevent this, wrap the formula in the `IFERROR` function:
> `=IFERROR((A2-A1)/A1, 0)`

#### Year-over-Year (YoY) Change

This compares a period to the same period in the previous year, which is crucial for seasonal data as it removes month-to-month volatility.

**Formula (for simple monthly data):**
```excel
=A13-A1
```
*(This assumes your data is listed monthly, so the value 12 rows above is the same month last year)*

**How it works:**
The formula subtracts the value from the corresponding prior year period from the current period's value.

> [!TIP]
> This simple offset is not robust. For more complex or unsorted data, a lookup formula is much safer. For example, to find the YoY change for a date in `B13`:
> `=A13 - XLOOKUP(EDATE(B13,-12), B:B, A:A)`

#### Dynamic Array Method (Excel 365)

This modern approach calculates the differences for an entire column with a single, dynamic formula, eliminating the need to drag anything down.

**Formula:**
```excel
=DROP(A2:A101, 1) - DROP(A2:A101, -1)
```

**How it works:**
This elegant formula performs array subtraction:
1.  `DROP(A2:A101, 1)`: This creates a new array from your data range but **drops the first value**. The resulting array is `{A3, A4, ..., A101}`.
2.  `DROP(A2:A101, -1)`: This creates another array but **drops the last value**. The resulting array is `{A2, A3, ..., A100}`.
3.  The formula then subtracts the second array from the first, element by element:
    *   The first calculation is `A3 - A2`.
    *   The second is `A4 - A3`, and so on.

This single formula spills the entire list of period-over-period differences automatically.
