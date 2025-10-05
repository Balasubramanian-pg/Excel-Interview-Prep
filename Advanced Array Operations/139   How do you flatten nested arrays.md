### 139. How do you flatten nested arrays?

Flattening an array means converting a two-dimensional range (with multiple rows and columns) into a single, one-dimensional list (a single column or row). This is extremely useful for creating lists for data validation, drop-downs, or further calculations. Modern Excel provides dedicated functions for this task.

#### Method 1: Flatten a Single Range with `TOCOL`

The `TOCOL` function is designed to transform an array or range into a single column. Its most powerful feature is the ability to automatically handle and ignore empty cells or errors.

**Formula:**
```excel
=TOCOL(A1:E10, 1)
```

**How it works:**
*   `A1:E10`: This is the two-dimensional source array you want to flatten.
*   `1`: This is the optional `[ignore]` argument, which tells the function what to skip.
    *   `0` or omitted: Keep all values (default).
    *   **`1`**: Ignore blank cells.
    *   `2`: Ignore errors.
    *   `3`: Ignore both blanks and errors.

The formula reads the range `A1:E10` row by row (by default) and stacks all the values into a single vertical list, skipping any cells that are empty.

> [!NOTE]
> The functions `TOCOL`, `TOROW`, `VSTACK`, and `HSTACK` are currently available only in Excel 365, Excel 2021, and Excel for the web.

#### Method 2: Flatten Multiple Non-Contiguous Ranges

Often, your data isn't in a single, neat block. You can combine multiple non-contiguous ranges into one list by first stacking them with `VSTACK` and then flattening the result.

**Formula:**
```excel
=TOCOL(VSTACK(A1:A10, C1:C10, E1:E10), 1)
```

**How it works:**
1.  `VSTACK(A1:A10, C1:C10, E1:E10)`: The `VSTACK` function takes the specified ranges (`A1:A10`, `C1:C10`, `E1:E10`) and stacks them vertically on top of each other, creating a single, multi-row array.
2.  `TOCOL(..., 1)`: The `TOCOL` function then takes this combined array and ensures it is returned as a clean, single column, ignoring any blank cells in the source ranges.

> [!TIP]
> **Creating a Unique, Sorted List**
> A very common requirement is to flatten a range and then get a sorted list of unique values. You can easily achieve this by wrapping the `TOCOL` function with `SORT` and `UNIQUE`.
>
> ```excel
> =SORT(UNIQUE(TOCOL(A1:E10, 1)))
> ```
> This single formula finds all non-blank values, removes duplicates, and sorts the final list alphabetically or numerically.
