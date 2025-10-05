### 62. **How do you find the nth occurrence of a value?**

Array formula:
=INDEX($A$1:$A$100, SMALL(IF($A$1:$A$100="SearchValue", ROW($A$1:$A$100)-ROW($A$1)+1), n))

Where n is the occurrence number (2 for second occurrence)

**Excel 365 alternative:**
=FILTER(A:A, A:A="SearchValue")
Returns all occurrences
