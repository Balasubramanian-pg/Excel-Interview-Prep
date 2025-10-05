### 56. **How do you remove duplicates with formulas?**

**Excel 365:**
=UNIQUE(A1:A100)

**Older Excel (array formula):**
=INDEX($A$1:$A$100, MATCH(0, COUNTIF($B$1:B1, $A$1:$A$100), 0))
Drag down, skips duplicates
