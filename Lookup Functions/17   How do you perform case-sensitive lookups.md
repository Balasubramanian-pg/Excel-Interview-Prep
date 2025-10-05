### 17. **How do you perform case-sensitive lookups?**

VLOOKUP is not case-sensitive. Use array formula:
=INDEX(return_range, MATCH(TRUE, EXACT(lookup_value, lookup_range), 0))

Example: =INDEX(B:B, MATCH(TRUE, EXACT("Apple", A:A), 0))
