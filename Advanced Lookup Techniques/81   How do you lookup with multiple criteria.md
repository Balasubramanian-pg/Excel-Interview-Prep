### 81. **How do you lookup with multiple criteria?**

**Method 1 - Helper column:**
Concatenate criteria: =A1&B1&C1
Then VLOOKUP on concatenated column

**Method 2 - INDEX-MATCH with arrays:**
=INDEX(ReturnRange, MATCH(1, (Criteria1Range=Criteria1)*(Criteria2Range=Criteria2), 0))
Array formula (Ctrl+Shift+Enter in older Excel)

**Method 3 - Excel 365 FILTER:**
=FILTER(Data, (Range1=Criteria1)*(Range2=Criteria2))
