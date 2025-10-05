### 216. **How do you use LAMBDA for custom functions?**

**Named formula - Custom discount:**

```
Discount = LAMBDA(amount, tier,
  amount * CHOOSE(tier, 0, 0.05, 0.10, 0.15, 0.20)
)

```

Use: =Discount(A1, B1)

**Recursive factorial:**

```
Factorial = LAMBDA(n,
  IF(n<=1, 1, n * Factorial(n-1))
)

```

**Complex business rule:**

```
PricingRule = LAMBDA(qty, customer_type, season,
  LET(
    base, 100,
    vol_discount, IF(qty>=100, 0.15, IF(qty>=50, 0.10, 0)),
    cust_discount, CHOOSE(customer_type, 0, 0.05, 0.10),
    seasonal, IF(season="Winter", 0.90, 1),
    base * (1-vol_discount) * (1-cust_discount) * seasonal
  )
)

```

**String processing:**

```
ExtractNumbers = LAMBDA(text,
  VALUE(CONCAT(IF(ISNUMBER(--MID(text, SEQUENCE(LEN(text)), 1)),
    MID(text, SEQUENCE(LEN(text)), 1), "")))
)

```
