### 136. How do you create cross-tabulation (pivot-like) formulas?

Creating a cross-tabulation (or "crosstab") with formulas allows you to build a dynamic summary table that updates automatically as your source data changes, without needing to manually refresh a PivotTable.

Here are three methods, ranging from the classic approach compatible with all Excel versions to the newest, most efficient functions available in Excel 365.

### Method 1: The Classic `SUMIFS` Formula (Manual Setup)

This is the most common and universally compatible method. It requires you to manually set up the row and column headers for your summary table.

**How it works:** You write a single `SUMIFS` formula in the top-left cell of your summary table and use mixed cell references (e.g., `$G2` and `H$1`) so you can drag the formula across and down to fill the table.

**Example:**
Imagine your source data is in columns A, B, and D:
*   `A:A` = Region
*   `B:B` = Product Category
*   `D:D` = Sales Amount

Your summary table is set up with unique Regions in column G (starting at `G2`) and unique Product Categories in row 1 (starting at `H1`).

1.  **In cell H2, enter the following formula:**
    ```excel
    =SUMIFS($D:$D, $A:$A, $G2, $B:$B, H$1)
    ```

2.  **Breakdown of the formula:**
    *   `$D:$D`: The range to sum (Sales). It's locked because it never changes.
    *   `$A:$A`: The first criteria range (Region). It's also locked.
    *   `$G2`: The first criterion (e.g., "North"). The column `G` is locked, but the row `2` is relative, so it will change to `G3`, `G4`, etc., when you drag down.
    *   `$B:$B`: The second criteria range (Product Category). It's locked.
    *   `H$1`: The second criterion (e.g., "Electronics"). The row `1` is locked, but the column `H` is relative, so it will change to `I1`, `J1`, etc., when you drag across.

3.  **Drag the formula** from H2 across to fill the row, then drag the entire row down to fill the table.

> [!CAUTION]
> This method is not fully dynamic. If a new Region or Product Category is added to your source data, you must manually add it to your summary table's headers for it to be included in the calculation.

### Method 2: The Modern Dynamic Array Approach (Excel 365/2021+)

This method uses modern Excel functions like `UNIQUE` and `TRANSPOSE` to generate the headers dynamically and then uses a single `SUMIFS` formula that spills the results across the entire table.

**How it works:** You create dynamic headers for rows and columns. Then, a single formula in the top-left cell references these dynamic header arrays to calculate all results at once.

**Example:**
Using the same source data, now in a table named `SalesData`.
*   `SalesData[Region]`
*   `SalesData[Product Category]`
*   `SalesData[Sales Amount]`

1.  **Create Dynamic Row Headers:** In cell `G2`, enter:
    ```excel
    =UNIQUE(SalesData[Region])
    ```
    This will spill a list of all unique regions downwards.

2.  **Create Dynamic Column Headers:** In cell `H1`, enter:
    ```excel
    =TRANSPOSE(UNIQUE(SalesData[Product Category]))
    ```
    This spills a list of unique product categories horizontally across the row.

3.  **Create the Crosstab Formula:** In cell `H2`, enter the single formula:
    ```excel
    =SUMIFS(SalesData[Sales Amount], SalesData[Region], G2#, SalesData[Product Category], H1#)
    ```

> [!IMPORTANT]
> The hash symbols (`#`) after `G2#` and `H1#` are called **spill operators**. They tell Excel to use the entire spilled array from those cells as the criteria, which causes this single `SUMIFS` formula to calculate and spill the results for the entire table.

### Method 3: The New `PIVOT` Function (Latest Excel 365)

The `PIVOT` function is a new function currently rolling out to Excel 365 Insiders. It is designed specifically for this purpose and is the simplest and most powerful formula-based solution.

**How it works:** The `PIVOT` function takes your data ranges and creates a complete summary table in one step, just like a PivotTable but with a single formula.

**Example:**
Using the same `SalesData` table.

1.  **In any empty cell, enter the formula:**
    ```excel
    =PIVOT(SalesData[Region], SalesData[Product Category], SalesData[Sales Amount], SUM)
    ```

2.  **Breakdown of the formula:**
    *   `SalesData[Region]`: The field to use for the rows.
    *   `SalesData[Product Category]`: The field to use for the columns.
    *   `SalesData[Sales Amount]`: The values to aggregate.
    *   `SUM`: The function to use for aggregation (can also be `AVERAGE`, `COUNT`, etc.).

> [!NOTE]
> The `PIVOT` and `GROUPBY` functions are very new and may not be available in your version of Excel 365 yet. They represent the future of data aggregation formulas in Excel. `GROUPBY` is used for single-criteria summaries, while `PIVOT` is specifically for two-criteria cross-tabulations.

> [!TIP]
> **Don't Forget PivotTables!**
> While formulas are great for dynamic dashboards, a standard **PivotTable** (Insert > PivotTable) is often the fastest and easiest way to create a crosstab summary. It offers powerful features like sorting, filtering, and drill-down with just a few clicks, though it requires a manual refresh to update.
