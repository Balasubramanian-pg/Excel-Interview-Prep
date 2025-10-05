### 182. BI: How do you calculate variance analysis?

Variance analysis is a core component of Business Intelligence (BI) that involves comparing actual results against a plan, budget, or forecast to identify and understand differences. These formulas are fundamental for creating financial reports, performance dashboards, and management summaries in Excel.

Assume your data is structured with `Actual` values in column B, `Budget` values in column C, and a `Category` (e.g., "Revenue" or "Expenses") in column A.

---

#### 1. Absolute Variance

This is the simplest form of variance, showing the raw numerical difference between the actual and budgeted figures.

**Formula:**
```excel
=B2 - C2
```
(Where `B2` is Actual, `C2` is Budget)

**How it works:**
This formula subtracts the budget from the actual amount.
*   A **positive** result means the actual was higher than the budget.
*   A **negative** result means the actual was lower than the budget.

> [!NOTE]
> The absolute variance is useful for understanding the magnitude of the difference in monetary terms, but it lacks context. A $10,000 variance is significant for a small business but might be negligible for a large corporation.

---

#### 2. Percentage Variance

This calculation expresses the absolute variance as a percentage of the budget, providing crucial context and making it easier to compare performance across items of different sizes.

**Formula:**
```excel
=(B2 - C2) / C2
```

**How it works:**
It calculates the absolute variance and then divides it by the budget amount. This normalizes the variance into a standardized percentage. Remember to format the cell as a Percentage.

> [!CAUTION]
> This formula will return a `#DIV/0!` error if the budget amount is zero. To handle this, wrap your formula in the `IFERROR` function:
> `=IFERROR((B2 - C2) / C2, "N/A")`

---

#### 3. Favorable / Unfavorable Indicator

A variance isn't inherently good or bad; its impact depends on the context. For revenue, a positive variance (more income than budgeted) is favorable. For expenses, a negative variance (less cost than budgeted) is favorable.

**Formula:**
```excel
=IF(A2="Revenue", IF(B2>C2, "Favorable", "Unfavorable"), IF(B2<C2, "Favorable", "Unfavorable"))
```
(Where `A2` contains the category type, e.g., "Revenue" or "Expenses")

**How it works:**
This nested `IF` statement first checks the category in cell `A2`.
*   **If it's "Revenue"**: It then checks if `Actual > Budget`. If true, it returns "Favorable"; otherwise, it's "Unfavorable".
*   **If it's not "Revenue"** (implying it's an expense): It checks if `Actual < Budget`. If true, it returns "Favorable"; otherwise, it's "Unfavorable".

> [!TIP]
> Use this formula as a basis for **Conditional Formatting**. You can create rules to automatically color-code cells green for "Favorable" and red for "Unfavorable," instantly making your reports easier to read.

---

#### 4. Variance Explanation / Categorization

To help focus attention on what matters most, you can automatically categorize the magnitude of a variance using the `IFS` function.

**Formula:**
```excel
=IFS(ABS(D2)<0.05, "Minimal", ABS(D2)<0.10, "Moderate", ABS(D2)<0.20, "Significant", TRUE, "Critical")
```
(Where `D2` contains the Percentage Variance)

**How it works:**
*   `ABS(D2)`: The `ABS` function returns the absolute value of the percentage variance, allowing you to treat a -15% variance and a +15% variance with the same level of importance.
*   `IFS(...)`: The `IFS` function checks a series of conditions in order.
    *   It first checks if the absolute variance is less than 5% (`0.05`) and labels it "Minimal".
    *   If not, it checks if it's less than 10% and labels it "Moderate", and so on.
    *   `TRUE`: This final condition acts as a catch-all for any variance of 20% or greater, labeling it "Critical".

> [!IMPORTANT]
> The thresholds used in this formula (5%, 10%, 20%) are examples. In a real-world scenario, these materiality thresholds should be defined by the business and may vary depending on the specific account or department being analyzed.
