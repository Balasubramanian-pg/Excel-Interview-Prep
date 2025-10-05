### 13. **What is XLOOKUP and how is it different?**

XLOOKUP (Excel 365/2021+) is the modern replacement:
Syntax: =XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])

Example: =XLOOKUP(E2, A2:A100, C2:C100, "Not Found")

**Advantages:**

- Simpler syntax
- Default exact match (no FALSE needed)
- Can search any direction
- Built-in error handling
- Can search from last to first
- Can return multiple columns
