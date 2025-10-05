### 139. **How do you flatten nested arrays?**

**Excel 365:**
=TOCOL(A1:E10, 1)
Converts 2D range to single column, ignoring blanks

**Flatten multiple non-contiguous ranges:**
=TOCOL(VSTACK(A1:A10, C1:C10, E1:E10))
