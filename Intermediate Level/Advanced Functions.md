### Advanced Functions

**Explain VLOOKUP:**
Searches for a value in the first column of a table and returns a value from another column in the same row.
Syntax: =VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
Example: =VLOOKUP(A2, D1:F10, 3, FALSE) finds A2 in column D and returns the corresponding value from column F

**VLOOKUP vs HLOOKUP:**

- **VLOOKUP** searches vertically (down columns)
- **HLOOKUP** searches horizontally (across rows)
Both work the same way, just in different directions

**INDEX-MATCH:**
More flexible than VLOOKUP. INDEX returns a value from a specific position, MATCH finds the position.
Syntax: =INDEX(return_range, MATCH(lookup_value, lookup_range, 0))
Example: =INDEX(C1:C100, MATCH(A2, B1:B100, 0))

**Advantages over VLOOKUP:**

- Can look left (VLOOKUP can only look right)
- Doesn't break if you insert/delete columns
- Faster with large datasets
- Can return entire rows or columns

**Nested IF statements:**
Multiple IF functions inside each other for complex logic.
Example: =IF(A1>=90, "A", IF(A1>=80, "B", IF(A1>=70, "C", "F")))
Limitation: Maximum of 64 nested IFs (but becomes hard to read after 3-4 levels)

**SUMIF and COUNTIF:**

- **SUMIF(range, criteria, [sum_range])**: Sums cells that meet a condition
Example: =SUMIF(A1:A10, ">100", B1:B10) sums B values where A is greater than 100
- **COUNTIF(range, criteria)**: Counts cells meeting criteria
Example: =COUNTIF(A1:A10, "Complete") counts cells containing "Complete"
