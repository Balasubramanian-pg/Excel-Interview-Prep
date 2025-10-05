### 123. How do you handle multiple conditions with scoring?

A weighted scoring system is a powerful technique for evaluating an item (like a sales lead, a project risk, or a candidate) based on multiple criteria, where each criterion contributes a different number of points to a total score. This is easily achieved by taking advantage of how Excel handles TRUE/FALSE logic in mathematical operations.

#### Weighted Scoring System using Boolean Logic

The most direct method is to write a simple sum where each condition is multiplied by its corresponding weight.

**Formula:**
```excel
=(A1="Yes")*10 + (B1>100)*20 + (C1="Premium")*15 + (D1>=EOMONTH(TODAY(),-1))*5
```

**How it works:**
This formula relies on a core Excel principle: when used in a calculation, `TRUE` is treated as `1` and `FALSE` is treated as `0`.

1.  Each condition is evaluated independently. For example, `(A1="Yes")` returns either `TRUE` or `FALSE`.
2.  The result is then multiplied by its weight.
    *   If `A1="Yes"` is `TRUE`, the calculation becomes `1 * 10 = 10` points.
    *   If `A1="Yes"` is `FALSE`, the calculation becomes `0 * 10 = 0` points.
3.  The `+` operators then sum the points from all the conditions that were met. If a condition is false, it simply adds zero to the total, having no effect on the score.

> [!IMPORTANT]
> The foundation of this technique is that Excel coerces Boolean values (`TRUE`/`FALSE`) into integers (`1`/`0`) during arithmetic operations. This allows you to "turn on" or "turn off" parts of a formula based on logical tests.

> [!NOTE]
> The user's provided example used `SUMPRODUCT`. While a powerful function, for a single-row scoring calculation like this, a simple sum using `+` is more direct, readable, and efficient. The `SUMPRODUCT` function is typically used when you need to perform this logic over an entire range of data in a single formula.

#### Making Complex Scores More Readable with `LET`

For scoring models with many conditions, the formula can become long and difficult to read. The `LET` function in Excel 365 allows you to assign names to your conditions, making the final calculation much cleaner.

**Formula:**
```excel
=LET(
    is_active, A1="Yes",
    high_value, B1>100,
    is_premium, C1="Premium",
    is_active*10 + high_value*20 + is_premium*15
)
```
This structure makes it far easier to see what each condition represents and to adjust the weights (`10`, `20`, `15`) without getting lost in a long formula.
