### 198. **How do you handle multi-level hierarchical aggregations?**

**Parent-Child Relationship Sum:**
=SUMIF(Parent_ID_Column, Current_ID, Value_Column) + Current_Row_Value

**Recursive hierarchy level:**
=IF(ISBLANK(XLOOKUP(A2, Parent_Col, Parent_Col)), 1,
1 + XLOOKUP(XLOOKUP(A2, ID_Col, Parent_Col), ID_Col, Level_Col))

**Path from root to node:**
=TEXTJOIN(" > ", TRUE,
XLOOKUP(A2, ID_Col, Name_Col),
XLOOKUP(XLOOKUP(A2, ID_Col, Parent_Col), ID_Col, Name_Col),
...
)

**Excel 365 - All descendants:**
=FILTER(ID_Col, ISNUMBER(SEARCH(Current_ID, Path_Col)))
