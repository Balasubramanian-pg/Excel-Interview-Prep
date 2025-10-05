### 14. **Explain XMATCH**

Partner to XLOOKUP, returns position:
Syntax: =XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])

Example: =XMATCH("Apple", A2:A100, 0) returns position

**match_mode:**

- 0: Exact match (default)
- 1: Exact match or next smallest
- 1: Exact match or next largest
- 2: Wildcard match
