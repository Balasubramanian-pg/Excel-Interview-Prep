### 143. **How do you audit complex formulas?**

**FORMULATEXT to document:**
=FORMULATEXT(A1)

**Create formula map:**
=SUBSTITUTE(FORMULATEXT(A1), ",", CHAR(10))
Shows formula with each argument on new line

**Trace precedents programmatically:**
No direct formula exists - use F2 (Edit) or Formulas → Trace Precedents
