### 61. **How do you find the last value in a column?**

**Method 1 - LOOKUP:**
=LOOKUP(2, 1/(A:A<>""), A:A)
Works because LOOKUP searches to the end

**Method 2 - INDEX-COUNTA:**
=INDEX(A:A, COUNTA(A:A))

**Method 3 - Excel 365:**
=FILTER(A:A, A:A<>"")
Then take the last value from results

**Method 4 - For numbers only:**
=LOOKUP(9.99E+307, A:A)
