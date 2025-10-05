### 119. How do you group data into bins/buckets?

Grouping data into bins or buckets is the process of converting continuous numerical data into a smaller number of discrete categories. This is essential for creating summaries, frequency distributions, and histograms.

### Method 1: The `FREQUENCY` Function (For Numerical Counts)

The `FREQUENCY` function is a specialized array function designed specifically to count how many values fall within specified numeric ranges.

**Formula:**
```excel
=FREQUENCY(Data_Range, Bins_Range)
```

**Example:**
Assume your data is in `A2:A101` and your bin thresholds are in `C2:C5` (containing 50, 100, 150, 200).
```excel
=FREQUENCY(A2:A101, C2:C5)
```

**How it works:**
*   `Data_Range`: The set of values you want to count.
*   `Bins_Range`: A range containing the **upper boundaries** of each bin. The values in this range must be in ascending order.
*   The function calculates counts for bins such as `<=50`, `>50 and <=100`, `>100 and <=150`, etc.

> [!IMPORTANT]
> The `FREQUENCY` function always returns an array with **one more element** than there are bins. This last element is a count of all values that are greater than the highest value in the `Bins_Range`. In the example above, it would count values >200.

> [!NOTE]
> In Excel 365, this formula spills the results automatically. In older versions, you must first select the range where the results will go (e.g., D2:D6) and enter the formula using **Ctrl+Shift+Enter**.

### Method 2: The `IFS` Function (For Categorical Labels)

This method is ideal when you want to assign a text label (like "Low" or "High") to each individual data point based on which bin it falls into.

**Formula:**
```excel
=IFS(A2<=50, "Low", A2<=100, "Medium", A2<=150, "High", TRUE, "Very High")
```

**How it works:**
`IFS` evaluates conditions in a specific sequence and stops as soon as it finds a `TRUE` condition.
1.  It first checks if `A2` is less than or equal to 50. If so, it returns "Low" and stops.
2.  If not, it checks if `A2` is less than or equal to 100, and so on.
3.  The `TRUE, "Very High"` pair at the end acts as a default or "catch-all" for any value that didn't meet the preceding criteria.

> [!CAUTION]
> The order of the conditions is critical. You must start with the smallest value and work your way up. If you checked for `<=150` first, a value of 40 would be incorrectly labeled as "High".

### Method 3: The Lookup Table Method (Most Flexible)

For complex or frequently changing bin definitions, using a lookup table is the most robust and maintainable method.

**Setup:**
1.  Create a small lookup table. The first column should contain the **lower boundary** of each bin, sorted in ascending order. The second column contains the corresponding category label.
    *   `D2`: `0`, `E2`: `Low`
    *   `D3`: `51`, `E3`: `Medium`
    *   `D4`: `101`, `E4`: `High`
    *   `D5`: `151`, `E5`: `Very High`

**Formula:**
```excel
=VLOOKUP(A2, $D$2:$E$5, 2, TRUE)
```

**How it works:**
*   The key is the final argument, `TRUE`, which tells `VLOOKUP` to perform an **approximate match**.
*   An approximate match finds the largest value in the first column of the lookup table that is less than or equal to your lookup value (`A2`). This behavior is exactly what is needed for binning.
*   For example, if `A2` contains `75`, `VLOOKUP` will match it with `51` and return the corresponding label, "Medium".

> [!TIP]
> This method separates your logic from your formula. If you need to change your bin thresholds or labels, you only have to update the lookup table, not every formula in your data column.
