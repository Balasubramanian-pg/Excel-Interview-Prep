### 12. **Explain INDEX and MATCH functions**

**INDEX(array, row_num, [col_num])**: Returns value at specific position
Example: =INDEX(C2:C100, 5) returns 5th value in column C

**MATCH(lookup_value, lookup_array, [match_type])**: Returns position of value

- match_type: 0 (exact), 1 (less than), -1 (greater than)
Example: =MATCH("Apple", A2:A100, 0) returns position of "Apple"

**Combined INDEX-MATCH:**
=INDEX(C2:C100, MATCH(E2, A2:A100, 0))
More powerful than VLOOKUP - can look left, doesn't break with column changes
