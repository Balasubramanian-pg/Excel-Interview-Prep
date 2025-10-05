### 83. **How do you lookup and return the last matching value?**

**Method 1 - Array formula:**
=LOOKUP(2, 1/(A:A=LookupValue), B:B)

**Method 2 - INDEX with aggregate:**
=INDEX(B:B, MAX(IF(A:A=LookupValue, ROW(A:A))))

**Excel 365:**
=INDEX(FILTER(B:B, A:A=LookupValue), COUNTA(FILTER(B:B, A:A=LookupValue)))
