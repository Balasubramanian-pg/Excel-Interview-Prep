### 223. **How do you debug complex formulas?**

**Formula evaluation steps with LET:**
=LET(
step1, A1*B1,
step1_debug, step1,  /* Can reference this to see intermediate result */
step2, step1/C1,
step2_debug, step2,
final, step2*D1,
final
)

**IFERROR with diagnostic:**
=IFERROR(
Complex_Formula,
"Error: " & ERROR.TYPE(Complex_Formula) & " at " & CELL("address")
)

**Trace formula dependencies:**
=FORMULATEXT(A1)
Then parse for cell references

**Test data validation:**
=IF(ISERROR(A1/B1), "Division error - check B1",
IF(B1=0, "B1 is zero",
A1/B1
)
)
