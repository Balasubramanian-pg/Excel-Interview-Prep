### 212. **How do you use REDUCE for cumulative operations?**

**Cumulative sum:**
=REDUCE(0, A1:A10, LAMBDA(acc, val, acc + val))

**Running maximum:**
=REDUCE(-9.99E+307, A1:A10, LAMBDA(acc, val, MAX(acc, val)))

**Cumulative product:**
=REDUCE(1, A1:A10, LAMBDA(acc, val, acc * val))

**Compound growth:**
=REDUCE(Initial_Value, Growth_Rates, LAMBDA(acc, rate, acc * (1 + rate)))

**String concatenation with separator:**
=REDUCE("", A1:A10, LAMBDA(acc, val, IF(acc="", val, acc & ", " & val)))
