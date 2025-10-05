### 58. **What is GETPIVOTDATA?**

Extracts data from PivotTable:
Syntax: =GETPIVOTDATA(data_field, pivot_table, [field1, item1], ...)

Example: =GETPIVOTDATA("Sales", $A$3, "Region", "West", "Product", "Widget")

**Advantages:**

- Reliable even if PivotTable layout changes
- Works with filtered PivotTables

**Disadvantages:**

- Verbose syntax
- Hard to copy across cells

**Tip:** Type = and click a PivotTable cell; Excel creates GETPIVOTDATA automatically
