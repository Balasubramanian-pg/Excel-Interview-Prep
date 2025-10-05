### 115. How do you split text into columns with formulas?

Splitting text from a single cell into multiple columns is a common data cleaning task. Using formulas for this makes the results dynamic, meaning they will automatically update if the source text changes.

### Method 1: The Modern `TEXTSPLIT` Function (Excel 365)

This is the simplest, most powerful, and recommended method for users with Excel 365. The `TEXTSPLIT` function can handle simple and complex splitting with a single command.

**To Split Text into Columns:**
This formula will take the text in cell A1 and spill it horizontally into adjacent columns, using a comma as the separator.
```excel
=TEXTSPLIT(A1, ",")
```

**To Split Text into Both Columns and Rows:**
You can specify different delimiters for columns and rows to transform a complex text string into a 2D array.
```excel
=TEXTSPLIT(A1, ",", ";")
```
*   In this example, commas (`,`) separate the values into different **columns**.
*   Semicolons (`;`) separate the values into different **rows**.

> [!NOTE]
> The `TEXTSPLIT` function spills the results automatically. You only need to enter the formula in one cell, and it will fill as many cells as needed. This single function replaces all of the complex legacy formulas.

### Method 2: Classic Formulas (for Excel 2019 and Earlier)

For older versions of Excel that do not have `TEXTSPLIT`, you must use a combination of text functions to manually extract each piece of the string. These formulas are significantly more complex and less flexible.

**Example:** Assume cell `A1` contains `First,Second,Last`

**1. Extract the First Item:**
```excel
=LEFT(A1, FIND(",", A1)-1)
```
*   `FIND(",", A1)` gets the position of the first comma.
*   `LEFT(...)` extracts all characters from the left up to that position, minus one to exclude the comma itself.
*   **Result:** `First`

**2. Extract a Middle Item (e.g., the second item):**
```excel
=MID(A1, FIND(",", A1)+1, FIND(",", A1, FIND(",", A1)+1) - FIND(",", A1)-1)
```
*   This complex formula calculates the start position and length needed to extract the text between the first and second commas.
*   **Result:** `Second`

**3. Extract the Last Item:**
```excel
=RIGHT(A1, LEN(A1) - FIND("*", SUBSTITUTE(A1, ",", "*", LEN(A1)-LEN(SUBSTITUTE(A1, ",", "")))))
```
*   This formula uses an advanced trick to find the position of the *last* delimiter and extracts everything to the right of it.
*   **Result:** `Last`

> [!WARNING]
> These legacy formulas are brittle. They can break if the number of delimiters changes or if a delimiter is not found, often returning `#VALUE!` errors. Each formula must be written and dragged separately for each column you want to create.

> [!TIP]
> **Don't Forget the "Text to Columns" Wizard**
> For one-time, non-dynamic splitting, the fastest method in any version of Excel is the **Text to Columns** wizard. Select your data, go to the **Data** tab, and click **Text to Columns**. This is a static operation—the results will not update if the original data changes.
