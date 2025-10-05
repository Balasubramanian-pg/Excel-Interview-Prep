### 213. **How do you use SCAN for running calculations?**

**Running total:**
=SCAN(0, A1:A10, LAMBDA(acc, val, acc + val))

**Running average:**
=SCAN(0, A1:A10, LAMBDA(acc, val,
LET(n, ROWS(OFFSET(A$1, 0, 0, ROW()-ROW(A$1)+1, 1)), (acc*(n-1) + val)/n)
))

**Fibonacci sequence:**
=SCAN({0,1}, SEQUENCE(20), LAMBDA(acc, n,
HSTACK(INDEX(acc,2), SUM(acc))
))

**Exponential smoothing:**
=SCAN(First_Value, Data, LAMBDA(acc, val, Alpha*val + (1-Alpha)*acc))

**Account balance tracker:**
=SCAN(Opening_Balance, Transactions, LAMBDA(acc, trans, acc + trans))
