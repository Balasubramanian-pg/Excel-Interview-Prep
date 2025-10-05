### 122. How do you create cascading conditions?

Cascading logic involves setting up a series of conditions that are evaluated in a specific order of priority. The formula stops and returns a value as soon as it finds the first condition that is true, ignoring all subsequent tests. The `IFS` function is perfectly designed for this type of sequential, priority-based logic.

#### Priority-Based Logic with `IFS`

The `IFS` function checks conditions one by one, making the order in which you write them critical.

**Formula:**
```excel
=IFS(C1="Override", "Special", A1>100, "High", B1="Priority", "Medium", TRUE, "Low")
```

**How it works:**
The formula evaluates each logical test in the sequence it is written. The first test to return `TRUE` determines the output, and the function goes no further.

1.  **Highest Priority Check:** First, it checks if `C1="Override"`. If `TRUE`, the formula immediately returns "Special" and stops all further evaluation.
2.  **Second Priority Check:** Only if the first test was `FALSE`, it moves on to check if `A1>100`. If this is `TRUE`, it returns "High" and stops.
3.  **Third Priority Check:** If the first two tests were `FALSE`, it checks if `B1="Priority"`. If `TRUE`, it returns "Medium" and stops.
4.  **Default / Catch-All:** If none of the preceding conditions were met, it evaluates the final test, `TRUE`. Since `TRUE` is always true, this pair acts as a default case, returning "Low".

> [!IMPORTANT]
> The order of your logical tests is crucial. `IFS` processes conditions sequentially and stops at the first `TRUE` result. Always place your most specific or highest-priority conditions at the beginning of the formula. For example, if `C1` was "Override" and `A1` was `150`, the result would be "Special", not "High", because the override check comes first.

> [!CAUTION]
> If no condition is met and you do not provide a final `TRUE` catch-all condition, the `IFS` function will return an `#N/A` error. Including `TRUE, "Default Value"` at the end is a best practice to handle all possibilities.

> [!TIP]
> The `SWITCH(TRUE, ...)` pattern, as shown in the previous topic, can also create cascading logic and is preferred by some users for its structure. The logic is identical: it evaluates each condition in order and returns the result for the first one that is `TRUE`.
