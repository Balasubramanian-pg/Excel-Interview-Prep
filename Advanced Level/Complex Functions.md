### Complex Functions

**XLOOKUP vs VLOOKUP:**
XLOOKUP (Excel 365/2021+) is the modern replacement:

- Syntax simpler: =XLOOKUP(lookup_value, lookup_array, return_array)
- Searches any direction (not just left-to-right)
- Default exact match (no need for FALSE)
- Can search from bottom-up
- Returns arrays for multiple matches
- Built-in error handling with 4th argument

**Array Formulas:**
Perform calculations on multiple values simultaneously and return multiple results. In older Excel, press Ctrl+Shift+Enter. In Excel 365, they're dynamic.
Example: =SUM(A1:A10*B1:B10) multiplies each pair and sums them

**OFFSET and INDIRECT:**

- **OFFSET(reference, rows, cols, [height], [width])**: Returns a reference offset from a starting cell. Dynamic and useful for moving ranges.
Example: =SUM(OFFSET(A1,0,0,5,1)) sums 5 cells starting from A1
- **INDIRECT(text_reference)**: Converts text to a cell reference
Example: =INDIRECT("A"&ROW()) creates dynamic cell references

**SUMIFS, COUNTIFS, AVERAGEIFS:**
Multiple criteria versions:

- =SUMIFS(sum_range, criteria_range1, criteria1, criteria_range2, criteria2...)
- =COUNTIFS(range1, criteria1, range2, criteria2...)
- =AVERAGEIFS(average_range, criteria_range1, criteria1...)
Example: =SUMIFS(D:D, A:A, "West", B:B, ">1000") sums column D where column A is "West" AND column B is greater than 1000

**TEXT Functions:**

- **CONCATENATE** or **&**: Joins text. =CONCATENATE(A1," ",B1) or =A1&" "&B1
- **LEFT(text, num_chars)**: Extracts characters from the left. =LEFT(A1,3) gets first 3 characters
- **RIGHT(text, num_chars)**: Extracts from the right
- **MID(text, start, num_chars)**: Extracts from the middle. =MID(A1,4,2) starts at position 4, takes 2 characters
