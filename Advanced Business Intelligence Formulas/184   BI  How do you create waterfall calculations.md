### 184. **BI: How do you create waterfall calculations?**

**Running Total for Waterfall:**
=SUM($B$2:B2)

**Floating Bar Start Position:**
=IF(B2>0, SUM($B$2:B1), SUM($B$2:B2))

**Floating Bar End Position:**
=SUM($B$2:B2)

**Excel 365 - Generate Waterfall Data:**
=LET(
values, A:A,
starts, SCAN(0, values, LAMBDA(acc, val, acc)),
ends, SCAN(0, values, LAMBDA(acc, val, acc+val)),
HSTACK(values, starts, ends)
)
