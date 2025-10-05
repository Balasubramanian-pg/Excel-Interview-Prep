### 118. How do you calculate cumulative percentages?

Calculating a cumulative or running total percentage is a common task in data analysis, especially for creating Pareto charts or understanding the contribution of items to a whole. The calculation involves dividing a running total by a grand total.

#### The Classic "Drag-Down" Formula

This method uses a mixed reference in the `SUM` function to create an expanding range for the running total, and an absolute reference for the grand total.

**Formula:**
```excel
=(SUM($B$2:B2) / SUM($B$2:$B$101))
```
*(Assuming your data is in B2:B101 with a header in B1)*

**How it works:**
1.  `SUM($B$2:B2)`: This is the **running total**. The `$` locks the starting cell of the sum (`B2`), but the ending cell (`B2`) is relative. As you drag this formula down to row 3, it becomes `SUM($B$2:B3)`, then `SUM($B$2:B4)`, and so on, creating a cumulative sum.
2.  `SUM($B$2:$B$101)`: This is the **grand total**. Both the start and end of the range are locked with `$` signs. This ensures that every running total is divided by the same, correct grand total for the entire dataset.

> [!IMPORTANT]
> After entering the formula, you must format the cell or column as a **Percentage** to display the result correctly (e.g., 0.25 will show as 25%).

---

### Application: Pareto Analysis (The 80/20 Rule)

The Pareto Principle states that for many events, roughly 80% of the effects come from 20% of the causes (e.g., 80% of revenue comes from 20% of customers). Cumulative percentages are essential for identifying this "vital few."

**Steps to Perform Pareto Analysis:**

1.  **Sort Data:** First, you must sort your data in **descending** order based on the value you are analyzing (e.g., sort by sales amount from largest to smallest). This is a critical step.

2.  **Calculate Individual Contribution:** (Optional but helpful) In a new column, calculate what percentage each individual item contributes to the total.
    `=B2/SUM($B$2:$B$101)`

3.  **Calculate Cumulative Percentage:** In the next column, use the cumulative percentage formula from above.

4.  **Analyze the Results:** Look down the cumulative percentage column. The rows where the value is 80% or less represent the top contributors (the "vital few"). By counting these items, you can see if the 80/20 rule applies to your data.

> [!TIP]
> Use Conditional Formatting to instantly highlight the rows that fall within the top 80%. Apply a rule to your cumulative percentage column with the formula `=C2<=0.8` (where C is the cumulative % column) to visually separate your top contributors.

---

### Dynamic Cumulative Percentage (Excel 365)

For a fully dynamic solution that spills all results from a single formula, you can use the `SCAN` function.

**Formula:**
```excel
=LET(
    data, B2:B101,
    running_total, SCAN(0, data, LAMBDA(acc, val, acc + val)),
    grand_total, SUM(data),
    running_total / grand_total
)
```
This single formula calculates the running total for the entire `data` range, divides it by the grand total, and spills the cumulative percentages for the entire list without needing to drag any formulas.
