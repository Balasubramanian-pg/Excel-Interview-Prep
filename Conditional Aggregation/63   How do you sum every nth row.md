# How to Sum Every nth Row in Excel

Sometimes you need to **sum values at fixed intervals**—for example, every 3rd, 5th, or 10th row—without adding helper columns. You can achieve this with a single formula using `SUMPRODUCT` and `MOD`.

## Formula Syntax

```excel
=SUMPRODUCT((MOD(ROW(A1:A100)-ROW(A1), n)=0)*(A1:A100))
```

Where:

* `A1:A100` → The range containing your data.
* `n` → The interval (for example, 3 for every 3rd row).
* `ROW(A1)` → Ensures the count starts from the first row of your selected range.

## Example: Sum Every 3rd Row

```excel
=SUMPRODUCT((MOD(ROW(A1:A100), 3)=0)*(A1:A100))
```

This adds the values in rows 3, 6, 9, 12, and so on.

> [!NOTE]
> The formula dynamically identifies row positions, so it continues to work even if you insert or delete rows later.

> [!TIP]
> Replace `A1:A100` with any column or named range to make the workbook easier to read.
> Example:
> `=SUMPRODUCT((MOD(ROW(Sales)-ROW(SalesFirst),3)=0)*(Sales))`

> [!IMPORTANT]
> The formula works correctly only on contiguous numeric ranges. Non-numeric values will be ignored.

> [!WARNING]
> Large ranges (like A1:A1000000) can slow performance because `SUMPRODUCT` evaluates each row individually.

> [!CAUTION]
> If your data starts below row 1, adjust `ROW(A1)` accordingly. Otherwise, the offset will misalign your calculation.

## Alternative: Using FILTER (Excel 365+)

If you’re using a modern version of Excel that supports dynamic arrays, this is a simpler and faster alternative:

```excel
=SUM(FILTER(A1:A100, MOD(ROW(A1:A100),3)=0))
```

This approach improves readability and performance on larger datasets.
