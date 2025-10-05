### 181. BI: How do you create rolling/moving averages?

Rolling or moving averages are a fundamental technique in Business Intelligence (BI) and data analysis used to smooth out short-term fluctuations in data and highlight longer-term trends. Here are several ways to calculate them in Excel, from classic formulas to modern dynamic arrays.

For these examples, assume your time-series data is in a column named `Data` (e.g., column A, starting in A2, with a header in A1) and you have defined a `Period` (e.g., 3 for a 3-day moving average).

#### Simple Moving Average (SMA)

An SMA calculates the average of a set of values over the last N periods. Each value in the period is weighted equally.

**Classic "Drag-Down" Formula:**
```excel
=AVERAGE(OFFSET(A2, COUNT($A$2:A2)-Period, 0, Period, 1))
```

**How it works:**
*   `COUNT($A$2:A2)`: This creates an expanding count as you drag the formula down (1, 2, 3, etc.).
*   `OFFSET(A2, ...)`: This function creates a dynamic range. It starts at `A2` and moves up by a number of rows equal to `COUNT(...) - Period` to find the start of the averaging window. It then creates a range that is `Period` rows high and 1 column wide.
*   `AVERAGE(...)`: Averages the values in the dynamic range created by `OFFSET`.

> [!CAUTION]
> The `OFFSET` function is **volatile**. This means it recalculates whenever *any* cell in the workbook changes, which can significantly slow down large spreadsheets. Use it with caution.

---

#### Weighted Moving Average (WMA)

A WMA gives more weight to recent data points, making it more responsive to new information than an SMA.

**Formula:**
```excel
=SUMPRODUCT(OFFSET(A2, COUNT($A$2:A2)-Period, 0, Period, 1), Weights) / SUM(Weights)
```
> [!IMPORTANT]
> For this formula, you must first define a range named `Weights` that contains your weighting factors (e.g., for a 3-period WMA, your weights could be `1, 2, 3` in cells `C1:C3` to give the most recent data 3x the weight of the oldest). The size of the `Weights` range must match the `Period`.

**How it works:**
*   The `OFFSET` part works identically to the SMA, selecting the data for the period.
*   `SUMPRODUCT(...)`: Multiplies each value in the data window by its corresponding value in the `Weights` range and sums the results.
*   `/ SUM(Weights)`: Normalizes the result by dividing by the sum of the weights.

---

#### Exponential Moving Average (EMA)

An EMA is a type of WMA that applies exponentially decreasing weights to older observations. It is highly responsive to recent changes and is widely used in finance and trend analysis.

**Recursive Formula:**
```excel
=IF(ROW()=2, A2, A2*Smoothing + B1*(1-Smoothing))
```
*   This formula assumes your data starts in `A2` and your EMA calculation starts in `B2`. `B1` refers to the previous EMA value.

**Setup:**
1.  Define the `Smoothing` constant. The standard formula is `2 / (Period + 1)`. You can put this value in a cell and reference it.
2.  The **first** EMA value has no preceding value. It is often seeded with the first data point (as in the formula above) or an SMA of the first few points.
3.  Enter the `IF` formula in the second row of your calculation column (`B2`) and drag it down. `B1` will automatically update to `B2`, `B3`, etc., creating the recursive link.

---

#### Excel 365 Dynamic Array Method (SMA)

This modern approach uses a single, non-volatile formula that spills all the results at once.

**Formula:**
```excel
=BYROW(SEQUENCE(ROWS(Data)-Period+1), LAMBDA(r, AVERAGE(INDEX(Data, r):INDEX(Data, r+Period-1))))
```

**How it works:**
*   `ROWS(Data)-Period+1`: Calculates the total number of moving averages to create.
*   `SEQUENCE(...)`: Generates a vertical array of numbers representing the starting row for each calculation window (e.g., 1, 2, 3...).
*   `BYROW(..., LAMBDA(r, ...))`: Iterates through each number (`r`) in the sequence generated above.
*   `LAMBDA(r, AVERAGE(...))`: For each `r`, it applies an `AVERAGE` function.
*   `INDEX(Data, r):INDEX(Data, r+Period-1)`: This is the core of the formula. It creates a "sliding" range reference for each iteration. For `r=1`, it averages rows 1 to 3. For `r=2`, it averages rows 2 to 4, and so on.

> [!TIP]
> This dynamic array method is the most efficient and robust solution in modern Excel. It avoids volatile functions, requires no dragging, and automatically adjusts if the source `Data` range changes size.
