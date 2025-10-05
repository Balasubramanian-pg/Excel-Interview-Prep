### 10. **Explain VLOOKUP in detail**

Syntax: =VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])

- **lookup_value**: What to search for
- **table_array**: Where to search (must include lookup column and return column)
- **col_index_num**: Which column number to return (1 is first column)
- **range_lookup**: FALSE/0 for exact match, TRUE/1 for approximate match

Example: =VLOOKUP(E2, A2:C100, 3, FALSE)
Looks for E2 in column A, returns value from column C

**Limitations:**

- Only looks to the right
- Breaks if columns are inserted/deleted
- Slower on large datasets
- Lookup column must be leftmost
