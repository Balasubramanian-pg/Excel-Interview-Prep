### 114. How do you unpivot data (columns to rows)?

Unpivoting is the process of transforming data from a "wide" format (with many columns) to a "long" format (with many rows). This is a crucial step for preparing data for analysis with tools like PivotTables, Power BI, or for use in database-style functions.

### Method 1: Power Query (Recommended Best Practice)

Power Query (found under the "Get & Transform Data" section of the Data tab) is by far the most powerful, reliable, and scalable way to unpivot data. It's a built-in tool designed specifically for this type of data transformation.

**How to use it:**
1.  Select your data range and go to the **Data** tab.
2.  Click **From Table/Range**. This will open the Power Query Editor.
3.  In the editor, select the columns you want to **keep as they are** (e.g., ID columns, name columns).
4.  Right-click on the header of any selected column and choose **Unpivot Other Columns**.
5.  Power Query will instantly transform all the other columns into two new ones: "Attribute" (the original column headers) and "Value".
6.  Click **Close & Load** to send the transformed data back to a new Excel sheet.

> [!IMPORTANT]
> Power Query is the industry-standard method for this task. It handles large datasets with ease, remembers your steps, and the transformation can be refreshed with a single click if your source data changes.

### Method 2: Manual Unpivot with Formulas (Excel 365)

For smaller, more contained datasets, you can manually unpivot using modern dynamic array functions like `VSTACK` and `HSTACK`. This approach requires you to build each segment of the unpivoted data and stack them together.

**Example:**
Imagine your data has `Product ID` in column A and sales for `Jan`, `Feb`, and `Mar` in columns B, C, and D.

**Formula:**
```excel
=VSTACK(
    HSTACK(A2:A10, "Jan", B2:B10),
    HSTACK(A2:A10, "Feb", C2:C10),
    HSTACK(A2:A10, "Mar", D2:D10)
)
```

**How it works:**
*   `HSTACK(A2:A10, "Jan", B2:B10)`: This creates the first block of the final table. It takes the `Product ID` column, adds a column with the hardcoded text "Jan", and then adds the corresponding sales values from the Jan column.
*   `VSTACK(...)`: This function then takes each of these horizontally stacked blocks and stacks them vertically on top of one another, creating the final long table.

> [!CAUTION]
> This formulaic approach is not dynamic. If you add a new month column (e.g., "Apr") to your source data, you must manually edit the formula to add a new `HSTACK` section for that month. It is best suited for tables with a fixed and small number of columns to unpivot.

### Clarification: Flattening vs. Unpivoting (`TOCOL`)

The `TOCOL` function is excellent for flattening a 2D range into a single column, but it is **not** a true unpivot function.

**Formula:**
```excel
=TOCOL(B2:D10, 1)
```
This formula will take all the sales values from the `Jan`, `Feb`, and `Mar` columns and stack them into one long list, ignoring any blanks.

> [!NOTE]
> The critical difference is that `TOCOL` **loses the context**. You get a list of values, but you no longer know which `Product ID` or which `Month` each value belongs to. True unpivoting preserves these relationships, which is why Power Query or the manual `HSTACK`/`VSTACK` method are the correct approaches.
