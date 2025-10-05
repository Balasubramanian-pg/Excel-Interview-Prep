### 65. **How do you sum based on partial text match?**

Use wildcard in SUMIF:
=SUMIF(A:A, "*West*", B:B)

Sums column B where column A contains "West"

**Multiple partial matches:**
=SUMIF(A:A, "*West*", B:B) + SUMIF(A:A, "*East*", B:B)
