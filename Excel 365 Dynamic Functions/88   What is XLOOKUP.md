### 88. **What is XLOOKUP?**

Modern replacement for VLOOKUP/HLOOKUP:
Syntax: =XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])

**Match modes:**

- 0: Exact match (default)
- 1: Exact or next smaller
- 1: Exact or next larger
- 2: Wildcard match

**Search modes:**

- 1: Search first to last (default)
- 1: Search last to first (reverse)
- 2: Binary search ascending
- 2: Binary search descending

Example: =XLOOKUP(A1, Names, Salaries, "Not Found", 0, -1)
