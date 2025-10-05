### 114. **How do you unpivot data (columns to rows)?**

**Excel 365 with TOCOL:**
=TOCOL(A2:E10, 1)
Converts all data to single column, ignoring blanks

**Stack with labels:**
=VSTACK(
HSTACK(A2:A10, "Col1", B2:B10),
HSTACK(A2:A10, "Col2", C2:C10)
)

**Best method:** Power Query (Get & Transform Data → Unpivot Columns)
