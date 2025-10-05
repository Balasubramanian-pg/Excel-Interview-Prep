### 70. **How do you count unique values with criteria?**

**Excel 365:**
=COUNTA(UNIQUE(FILTER(A:A, B:B="Criteria")))

**Older Excel (array formula):**
=SUM(IF(B1:B100="Criteria", 1/COUNTIFS(A1:A100, A1:A100, B1:B100, "Criteria"), 0))
