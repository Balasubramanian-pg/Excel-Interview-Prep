### 103. **How do you calculate moving/rolling averages?**

**Simple (drag down):**
=AVERAGE(A1:A10), =AVERAGE(A2:A11), etc.

**Dynamic with OFFSET:**
=AVERAGE(OFFSET(A1, ROW()-1, 0, 10, 1))

**Excel 365:**
=BYROW(SEQUENCE(COUNTA(A:A)-9), LAMBDA(r, AVERAGE(OFFSET(A1, r-1, 0, 10, 1))))
