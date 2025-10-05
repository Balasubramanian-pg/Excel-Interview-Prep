### 184. BI: How do you create waterfall calculations?

Waterfall charts are powerful BI tools for visualizing how a starting value is affected by a series of positive and negative changes, leading to a final value. Creating one requires structuring the data correctly to show the "floating" bars that represent these changes.

> [!NOTE]
> Modern Excel has a built-in **Waterfall chart type** (Insert > Chart > Waterfall). While this chart type automates much of the process, understanding the underlying calculations is crucial for customization and for creating the chart in older Excel versions.

#### Method 1: Classic Helper Column Approach

This traditional method uses helper columns to calculate the start and end positions for each floating bar, which are then plotted using a stacked bar chart. Assume your positive and negative change values are in column B, starting from B2.

1.  **Running Total / End Position:**
    In cell C2, calculate the cumulative sum. This value represents the top of each bar.
    ```excel
    =SUM($B$2:B2)
    ```
    The `$` locks the starting cell, so as you drag the formula down, the range expands to create a running total.

2.  **Floating Bar Start Position:**
    This formula calculates the "base" from which each bar should start.
    ```excel
    =IF(B2>0, SUM($B$2:B1), SUM($B$2:B2))
    ```

    > [!IMPORTANT]
    > This logic is the key to the waterfall effect:
    > *   If the change (`B2`) is **positive**, the bar should start at the *previous* running total (`SUM($B$2:B1)`).
    > *   If the change (`B2`) is **negative**, the bar should start at the *current* running total (`SUM($B$2:B2)`) so it can draw downwards.

To create the chart, you would typically need additional columns for the "invisible" base, the positive change ("Up" series), and the negative change ("Down" series) to build a custom stacked column chart.

#### Method 2: Excel 365 Dynamic Array Approach

The `LET` and `SCAN` functions in Excel 365 can generate the entire data structure for a waterfall chart with a single, dynamic formula.

**Formula:**
```excel
=LET(
    values, B2:B10,
    ends, SCAN(0, values, LAMBDA(acc, val, acc + val)),
    starts, ends - values,
    HSTACK(values, starts, ends)
)
```

**How it works:**
*   `LET(...)`: This function assigns names to calculations, making the formula cleaner and more efficient.
    *   `values, B2:B10`: First, we name our input range of positive/negative changes `values`.
    *   `ends, SCAN(...)`: We calculate the running total (the **end** point of each bar) using `SCAN`. The `LAMBDA` function `acc + val` simply adds the current value to the accumulated total for each row.
    *   `starts, ends - values`: This is an elegant way to find the **start** point of each bar. By subtracting the current change value from the end position, you get the running total from the *previous* step.
    *   `HSTACK(...)`: Finally, `HSTACK` combines the original `values`, the calculated `starts`, and the calculated `ends` into a single three-column array, ready for charting.

> [!TIP]
> This dynamic array formula is non-volatile and automatically updates if the `values` range changes. It provides all the necessary data in one place without needing to drag formulas down multiple helper columns.
