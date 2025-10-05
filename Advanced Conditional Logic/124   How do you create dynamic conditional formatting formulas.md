### 124. How do you create dynamic conditional formatting formulas?

Conditional formatting becomes incredibly powerful when you use custom formulas. A formula-based rule applies formatting to a cell whenever the formula you write evaluates to `TRUE`. The key to making these formulas dynamic is the correct use of absolute (`$A$1`), relative (`A1`), and mixed (`$A1` or `A$1`) cell references.

To apply these rules: Go to **Home > Conditional Formatting > New Rule > Use a formula to determine which cells to format**.

#### Highlight Entire Row Based on a Cell's Value

This is one of the most common tasks: color-coding an entire row (e.g., for a task, project, or order) based on its status in a single column.

**Formula:**
```excel
=$E1="Complete"
```
**Apply to:** `$A$1:$Z$1000`

**How it works:**
The formula is written from the perspective of the top-left cell of your "Applies to" range (A1). The dollar sign (`$`) is critical.
*   `$E1`: The `$` **locks the column** to E. This tells Excel that no matter which cell is being evaluated (A1, B1, C1...), it must always look at the value in column **E**.
*   The row number `1` is **relative**. When the rule is evaluated for the next row, it automatically adjusts to check `$E2`, then `$E3`, and so on.

#### Alternate Row Shading (Banding)

This makes large tables much easier to read by applying a light color to every other row.

**Formula:**
```excel
=MOD(ROW(),2)=0
```

**How it works:**
*   `ROW()`: This function returns the row number of the cell being evaluated.
*   `MOD(..., 2)`: This calculates the remainder when the row number is divided by 2. For even rows, the remainder is `0`; for odd rows, it is `1`.
*   `=0`: The condition is `TRUE` only for even-numbered rows, which get formatted.

> [!TIP]
> To shade odd rows instead, change the formula to `=MOD(ROW(),2)=1`.

#### Highlight Duplicates in a Column

This formula highlights the second, third, and subsequent occurrences of a value, leaving the first instance unformatted.

**Formula:**
```excel
=COUNTIF($A$1:$A1, $A1)>1
```
**Apply to:** `$A$1:$A$1000`

**How it works:**
This formula uses an "expanding range" to check for duplicates.
*   `$A$1:$A1`: The first part of the range is locked, but the second is relative.
*   As Excel evaluates down the column, the range grows: for cell A5, the formula checks `=COUNTIF($A$1:$A5, $A5)`.
*   It counts how many times the current cell's value has appeared *from the top of the column down to the current row*. The condition `>1` is only `TRUE` if the value has already appeared at least once before.

#### Highlight Dates Within the Next 7 Days

This is perfect for dashboards to draw attention to upcoming deadlines or events.

**Formula:**
```excel
=AND(A1>=TODAY(), A1<=TODAY()+7)
```

**How it works:**
*   `TODAY()`: A volatile function that always returns the current date.
*   `AND(...)`: This ensures two conditions are met: the date in `A1` is on or after today, AND the date in `A1` is on or before seven days from now.

#### Highlight the Top 10% of Values

This rule dynamically identifies the top-performing values in a range, even if the numbers change.

**Formula:**
```excel
=A1>=PERCENTILE($A$1:$A$100, 0.9)
```
**Apply to:** `$A$1:$A$100`

**How it works:**
*   `PERCENTILE($A$1:$A$100, 0.9)`: This calculates the value at the 90th percentile for the entire range. This is the minimum value an item needs to have to be in the top 10%.
*   `A1>=...`: The formula then checks if the value of the current cell (`A1`) is greater than or equal to that threshold.
*   The range inside `PERCENTILE` must be absolute (`$A$1:$A$100`) so that every cell is compared against the same, correct threshold.
