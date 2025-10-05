### 35. **How do you sum with multiple OR criteria?**

Use multiple SUMIF functions:
=SUMIF(A:A, "West", B:B) + SUMIF(A:A, "East", B:B)

Or use SUMPRODUCT:
=SUMPRODUCT((A:A="West")+(A:A="East"), B:B)
