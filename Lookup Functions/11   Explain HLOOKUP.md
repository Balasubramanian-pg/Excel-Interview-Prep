### 11. **Explain HLOOKUP**

Same as VLOOKUP but horizontal:
Syntax: =HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])

Example: =HLOOKUP("Sales", A1:F5, 3, FALSE)
Looks for "Sales" in row 1, returns value from row 3
