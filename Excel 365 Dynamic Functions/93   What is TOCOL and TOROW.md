### 93. **What is TOCOL and TOROW?**

Convert arrays to single column/row:

**TOCOL(array, [ignore], [scan_by_column]):**
=TOCOL(A1:C10) converts 3-column range to single column

**TOROW(array, [ignore], [scan_by_column]):**
=TOROW(A1:A10) converts column to row

**Ignore parameter:**

- 0: Keep all (default)
- 1: Ignore blanks
- 2: Ignore errors
- 3: Ignore both
