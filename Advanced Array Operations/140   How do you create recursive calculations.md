### 140. How do you create recursive calculations?

Recursive calculations, where a function calls itself to solve a problem, are possible in Excel through the `LAMBDA` function. This advanced feature allows you to create custom, self-referencing functions for complex problems like calculating factorials or Fibonacci sequences without needing helper columns.

> [!IMPORTANT]
> A `LAMBDA` function can only perform recursion when it is defined as a **Named Range** in the Name Manager. You cannot create a recursive function by typing it directly into a cell.

#### How to Create a Recursive `LAMBDA`

1.  Go to the **Formulas** tab and click on **Name Manager**.
2.  Click **New...**.
3.  In the **Name** field, enter the name you will use to call the function (e.g., `Fibonacci`).
4.  In the **Refers to** field, enter your `LAMBDA` formula.
5.  Click **OK**.

---

#### Example 1: Fibonacci Sequence

The Fibonacci sequence is a classic example of recursion, where each number is the sum of the two preceding ones (e.g., 0, 1, 1, 2, 3, 5, 8...).

**Named Function Definition:**
*   Name: `Fib`
*   Refers to:
    ```excel
    =LAMBDA(n, IF(n<=1, n, Fib(n-1) + Fib(n-2)))
    ```

**How it works:**
*   `LAMBDA(n, ...)`: Defines a function that accepts one argument, `n` (the position in the sequence).
*   `IF(n<=1, n, ...)`: This is the **base case** or stopping condition. If `n` is 0 or 1, the function simply returns `n`. This is crucial to prevent an infinite loop.
*   `... Fib(n-1) + Fib(n-2)`: This is the **recursive step**. For any `n` greater than 1, the function calls itself twice: once for the previous number (`Fib(n-1)`) and once for the number before that (`Fib(n-2)`), and adds their results.

**Usage in a cell:**
To find the 10th Fibonacci number, you would enter:
```excel
=Fib(10)
```

---

#### Example 2: Factorial

The factorial of a number is the product of all positive integers up to that number (e.g., 5! = 5 × 4 × 3 × 2 × 1 = 120).

**Named Function Definition:**
*   Name: `Factorial`
*   Refers to:
    ```excel
    =LAMBDA(n, IF(n<=1, 1, n * Factorial(n-1)))
    ```

**How it works:**
*   `LAMBDA(n, ...)`: Defines a function that accepts one argument, `n`.
*   `IF(n<=1, 1, ...)`: The **base case**. The factorial of 0 or 1 is defined as 1.
*   `... n * Factorial(n-1)`: The **recursive step**. The function multiplies `n` by the result of calling itself with a smaller argument (`n-1`).

**Usage in a cell:**
To calculate the factorial of 5, you would enter:
```excel
=Factorial(5)
```

> [!WARNING]
> Recursive functions can be computationally expensive and may cause Excel to become slow or unresponsive if used with large input numbers. The Fibonacci example provided is particularly inefficient and is intended for demonstration purposes.
