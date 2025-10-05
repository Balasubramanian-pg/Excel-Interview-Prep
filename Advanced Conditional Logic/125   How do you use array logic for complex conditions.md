### 125. How do you use array logic for complex conditions?

Array logic, traditionally performed with the `SUMPRODUCT` function, allows you to apply complex criteria across entire ranges of data in a single formula. This technique relies on Excel's ability to treat `TRUE` and `FALSE` values as `1` and `0` in mathematical calculations, enabling you to build powerful conditional counting and summing formulas.

> [!NOTE]
> While modern functions like `COUNTIFS`, `SUMIFS`, and `FILTER` are often more efficient for these tasks, understanding `SUMPRODUCT` is key to mastering advanced array logic, especially for scenarios that `...IFS` functions cannot handle (like complex `OR` logic combined with sums).

#### The Core Principle: `AND` vs. `OR` Logic

*   **`AND` Logic (Multiplication `*`)**: When you multiply arrays of `TRUE`/`FALSE` values, the result is `1` only if **all** conditions in that row are `TRUE` (`1 * 1 * 1 = 1`). If any condition is `FALSE`, the result becomes `0` (`1 * 0 * 1 = 0`).
*   **`OR` Logic (Addition `+`)**: When you add arrays of `TRUE`/`FALSE` values, the result is `1` or more if **any** condition in that row is `TRUE` (`1 + 0 = 1`). The result is `0` only if all conditions are `FALSE` (`0 + 0 = 0`).

---

#### Counting Rows with Multiple `AND` Conditions

This is the classic use case for `SUMPRODUCT`, where you count rows that meet several criteria simultaneously.

**Example: Count active records for Product X with sales between two dates.**
```excel
=SUMPRODUCT((Product_Column="X") * (Status_Column="Active") * (Date_Column>=StartDate) * (Date_Column<=EndDate))
```

**How it works:**
1.  `(Product_Column="X")`: Creates an array of `TRUE`/`FALSE` values (e.g., `{TRUE; FALSE; TRUE}`).
2.  Each subsequent condition creates its own `TRUE`/`FALSE` array.
3.  The `*` operator multiplies these arrays element by element. A row's final result is `1` only if all conditions for that row were `TRUE`.
4.  `SUMPRODUCT` then sums the final array of `1`s and `0`s, giving you the total count of rows that matched all criteria.

> [!IMPORTANT]
> For standard `AND` conditions, using `COUNTIFS` or `SUMIFS` is highly recommended. They are significantly faster on large datasets because they are specifically optimized for this task.
>
> **Modern Equivalent:**
> `=COUNTIFS(Product_Column, "X", Status_Column, "Active", Date_Column, ">="&StartDate, ...)`

---

#### Summing with Complex `OR` Logic

This is where `SUMPRODUCT` truly shines, as `SUMIFS` cannot easily handle `OR` conditions on the same column combined with other criteria.

**Example: Sum the sales (Column B) where the region (Column A) is either "West" OR "East".**
```excel
=SUMPRODUCT(((A1:A100="West") + (A1:A100="East") > 0) * B1:B100)
```

**How it works:**
1.  `(A1:A100="West")`: Creates a `TRUE`/`FALSE` array.
2.  `(A1:A100="East")`: Creates a second `TRUE`/`FALSE` array.
3.  `(...) + (...)`: The addition operator acts as a logical `OR`. If a cell in column A is "West" or "East", the result for that row will be `1` (`1+0` or `0+1`). If it's neither, the result is `0`.
4.  `>0`: This is a critical step that converts any result greater than `0` (e.g., `1` from a single match, or `2` if a cell could match both) back into a simple `TRUE` (`1`). This creates a final `TRUE`/`FALSE` array indicating which rows to include.
5.  `* B1:B100`: This final criteria array (`{1;0;1;...}`) is multiplied by the corresponding values in the sales column (`B1:B100`), effectively zeroing out the sales for rows that didn't match.
6.  `SUMPRODUCT` sums the resulting array of values.

> [!TIP]
> A modern and very clean way to handle this `OR` logic is to provide an array constant `{}` directly to `SUMIFS` and wrap it in `SUM`.
>
> **Modern Equivalent:**
> `=SUM(SUMIFS(B1:B100, A1:A100, {"West","East"}))`
