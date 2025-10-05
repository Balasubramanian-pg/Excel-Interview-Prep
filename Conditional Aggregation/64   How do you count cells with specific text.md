### 64. **How do you count cells with specific text?**

**Exact match:**
=COUNTIF(A:A, "Apple")

**Contains text (wildcard):**
=COUNTIF(A:A, "*apple*")

**Starts with:**
=COUNTIF(A:A, "apple*")

**Ends with:**
=COUNTIF(A:A, "*apple")

**Case-sensitive count:**
=SUMPRODUCT(--EXACT(A1:A100, "Apple"))
