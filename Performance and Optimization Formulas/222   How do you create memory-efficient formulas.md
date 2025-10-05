### 222. **How do you create memory-efficient formulas?**

**Instead of entire column references:**
Bad: =SUMIF(A:A, "X", B:B)
Good: =SUMIF(A1:A1000, "X", B1:B1000)

**Use Tables for auto-expanding ranges:**
=SUMIF(Table[Category], "X", Table[Amount])

**Consolidate repeated calculations with LET:**
=LET(
calc, EXPENSIVE_CALCULATION(A1:A1000),
calc * 2 + calc / 3
)

**Avoid array formulas in conditional formatting when possible**
