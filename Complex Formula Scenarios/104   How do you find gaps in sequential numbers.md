### 104. **How do you find gaps in sequential numbers?**

**Missing numbers:**
=FILTER(SEQUENCE(MAX(A:A)), ISNA(XMATCH(SEQUENCE(MAX(A:A)), A:A)))

Returns all missing numbers in sequence

**Older Excel:** Helper column with:
=IF(COUNTIF($A$1:$A$100, ROW())=0, ROW(), "")
