### 96. **What is TAKE and DROP?**

Extract portions from arrays:

**TAKE(array, rows, [cols]):**

- Positive: Take from start
- Negative: Take from end
=TAKE(A1:A100, 10) returns first 10 rows
=TAKE(A1:A100, -5) returns last 5 rows

**DROP(array, rows, [cols]):**
=DROP(A1:A100, 1) removes header row
=DROP(A1:A100, -5) removes last 5 rows
