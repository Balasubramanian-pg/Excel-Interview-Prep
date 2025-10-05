### 68. **How do you sum with OR conditions?**

**Method 1 - Multiple SUMIF:**
=SUMIF(A:A, "West", B:B) + SUMIF(A:A, "East", B:B)

**Method 2 - SUMPRODUCT:**
=SUMPRODUCT((A:A="West")+(A:A="East"), B:B)

**Method 3 - Array formula:**
=SUM(IF((A:A="West")+(A:A="East"), B:B, 0))
