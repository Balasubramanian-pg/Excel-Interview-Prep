### 121. How do you create complex nested conditions?

Handling multiple conditions in Excel can quickly lead to long, convoluted nested `IF` statements that are difficult to write and debug. Modern Excel functions like `SWITCH` and `IFS` provide much cleaner and more readable alternatives.

#### Method 1: Using `SWITCH` for a Single Variable

The `SWITCH` function is the ideal replacement for a nested `IF` when you are checking a single cell or expression against a list of possible exact values.

**Formula:**
```excel
=SWITCH(A1, "A", "Excellent", "B", "Good", "C", "Average", "D", "Poor", "F", "Fail", "Invalid Grade")
```

**How it works:**
The function evaluates the expression (in this case, the value in cell `A1`) and compares it against a series of `value/result` pairs.
*   `A1`: The expression to evaluate.
*   `"A", "Excellent"`: If `A1` is "A", return "Excellent".
*   `"B", "Good"`: If `A1` is "B", return "Good".
*   ...and so on.
*   `"Invalid Grade"`: This is the optional `default` value. If `A1` does not match any of the specified values, this result is returned.

> [!NOTE]
> This single `SWITCH` function is equivalent to the following nested `IF` statement, but is significantly easier to read and manage:
> `=IF(A1="A", "Excellent", IF(A1="B", "Good", IF(A1="C", "Average", ...)))`

#### Method 2: Using `SWITCH(TRUE, ...)` for Complex Logical Conditions

You can use a clever variation of `SWITCH` to handle multiple, complex conditions involving different variables or logical tests (e.g., `AND`, `OR`). This pattern provides a powerful alternative to the `IFS` function.

**Formula:**
```excel
=SWITCH(TRUE, AND(A1>90, B1="Y"), "Tier 1", AND(A1>80, B1="Y"), "Tier 2", AND(A1>70), "Tier 3", "No Tier")
```

**How it works:**
*   `SWITCH(TRUE, ...)`: The expression being evaluated is `TRUE`. The function then scans the list of conditions, looking for the **first one** that evaluates to `TRUE`.
*   `AND(A1>90, B1="Y"), "Tier 1"`: If the score in `A1` is greater than 90 AND the status in `B1` is "Y", the condition is `TRUE`. `SWITCH` finds this first true condition and returns "Tier 1".
*   `AND(A1>80, B1="Y"), "Tier 2"`: If the first condition was false, it checks this one next.
*   `"No Tier"`: This acts as the default value if none of the preceding conditions evaluate to `TRUE`.

> [!IMPORTANT]
> The order of your conditions is critical. `SWITCH` (and `IFS`) will stop at the first condition that returns `TRUE`. You should always list your most specific conditions first and your more general conditions later.

> [!TIP]
> The `IFS` function was designed specifically for this scenario and may feel more intuitive to some users. The equivalent `IFS` formula would be:
> `=IFS(AND(A1>90, B1="Y"), "Tier 1", AND(A1>80, B1="Y"), "Tier 2", A1>70, "Tier 3", TRUE, "No Tier")`
> Both `SWITCH(TRUE, ...)` and `IFS` are excellent, modern solutions for avoiding complex nested `IF` statements.
