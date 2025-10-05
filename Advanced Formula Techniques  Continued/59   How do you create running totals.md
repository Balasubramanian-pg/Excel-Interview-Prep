### 59. **How do you create running totals?**

**Method 1 - Simple:**
=SUM($A$1:A1) and drag down (expanding range)

**Method 2 - SUMIF for grouped data:**
=SUMIF($A$1:A1, A1, $B$1:B1)

**Method 3 - Excel 365 SCAN:**
=SCAN(0, A1:A100, LAMBDA(acc, val, acc + val))
