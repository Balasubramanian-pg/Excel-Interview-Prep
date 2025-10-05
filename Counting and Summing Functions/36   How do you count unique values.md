### 36. **How do you count unique values?**

**Excel 365:** =COUNTA(UNIQUE(A1:A100))

**Older Excel (array formula):**
=SUMPRODUCT(1/COUNTIF(A1:A100, A1:A100))

Or: =SUM(1/COUNTIF(A1:A100, A1:A100))
