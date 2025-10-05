### 57. **How do you rank values? (Continued)**

- **RANK.AVG(number, ref, [order])**: Rank with ties getting average rank
Example: If two values tie for 3rd, both get 3.5
- **PERCENTRANK.INC(array, x, [significance])**: Rank as percentile
Example: =PERCENTRANK.INC($A$1:$A$100, A1, 3) returns percentile rank with 3 decimals

**Handle duplicates differently:**
=RANK.EQ(A1, $A$1:$A$100) + COUNTIF($A$1:A1, A1) - 1
This gives unique ranks even for duplicates
