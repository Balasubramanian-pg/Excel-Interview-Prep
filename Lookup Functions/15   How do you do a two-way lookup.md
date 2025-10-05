### 15. **How do you do a two-way lookup?**

Find value based on both row and column criteria:

**Method 1 - INDEX with two MATCH:**
=INDEX(data_range, MATCH(row_value, row_range, 0), MATCH(col_value, col_range, 0))

Example: =INDEX(B2:E10, MATCH("Product A", A2:A10, 0), MATCH("Q2", B1:E1, 0))

**Method 2 - XLOOKUP nested:**
=XLOOKUP(col_value, col_range, XLOOKUP(row_value, row_range, data_range))
