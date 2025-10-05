### 220. **How do you create self-referencing dynamic arrays?**

**Warning:** These can be tricky and potentially unstable

**Iterative calculation:**
=LET(
initial, A1,
iterations, 100,
REDUCE(initial, SEQUENCE(iterations),
LAMBDA(acc, n, acc * 0.9 + 10)
)
)

**Expanding sequences:**
=XLOOKUP(ROW(),
SEQUENCE(ROWS(Data)),
SCAN(First_Value, Data, LAMBDA(acc, val, acc + val))
)

**Conditional cumulative:**
=SCAN(0, A:A, LAMBDA(acc, val,
IF(val="Reset", 0, acc + val)
))
