### 78. **How do you count specific characters in text?**

=(LEN(A1)-LEN(SUBSTITUTE(A1, "a", "")))/LEN("a")

Counts occurrences of "a" in A1

**Count spaces:**
=LEN(A1)-LEN(SUBSTITUTE(A1, " ", ""))

**Count words:**
=LEN(TRIM(A1))-LEN(SUBSTITUTE(A1, " ", ""))+1
