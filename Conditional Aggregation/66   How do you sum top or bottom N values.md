### 66. **How do you sum top or bottom N values?**

**Top N:**
=SUMPRODUCT(LARGE(A1:A100, ROW(INDIRECT("1:"&N))))

**Bottom N:**
=SUMPRODUCT(SMALL(A1:A100, ROW(INDIRECT("1:"&N))))

**Example - top 5:**
=SUMPRODUCT(LARGE(A1:A100, {1;2;3;4;5}))

Or: =SUM(LARGE(A1:A100, ROW(1:5)))
