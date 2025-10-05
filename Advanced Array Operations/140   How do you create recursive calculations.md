### 140. **How do you create recursive calculations?**

**Fibonacci sequence:**
Named formula approach using LAMBDA and recursion:

```
Fib = LAMBDA(n, IF(n<=1, n, Fib(n-1) + Fib(n-2)))

```

**Factorial:**

```
Factorial = LAMBDA(n, IF(n<=1, 1, n * Factorial(n-1)))

```

**Note:** Must be saved as named functions, not directly in cells
