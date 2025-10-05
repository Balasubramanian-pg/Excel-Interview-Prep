### 85. **How do you lookup nearest value?**

**Closest match:**
=INDEX(ReturnRange, MATCH(MIN(ABS(LookupRange-LookupValue)), ABS(LookupRange-LookupValue), 0))

Array formula in older Excel

**Excel 365:**
=LET(diff, ABS(LookupRange-LookupValue), INDEX(ReturnRange, MATCH(MIN(diff), diff, 0)))
