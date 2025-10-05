### Practical Scenarios

**How to clean messy data:**

1. Use TRIM() to remove extra spaces
2. Use PROPER(), UPPER(), or LOWER() for consistent capitalization
3. Find & Replace for common errors
4. Text to Columns for splitting data
5. Remove duplicates
6. Data validation for future entries
7. Power Query for complex transformations

**Analyze sales data across regions/time:**

1. Create PivotTable with dates in Rows, regions in Columns
2. Group dates by months/quarters
3. Add sales values
4. Use slicers for interactive filtering
5. Create PivotChart for visualization
6. Use GETPIVOTDATA for specific values in formulas
7. Add calculated fields for metrics like growth rate

**Handle formula errors:**

- **IFERROR(formula, value_if_error)**: Returns specified value if formula errors
Example: =IFERROR(VLOOKUP(A1,D:E,2,0),"Not Found")
- **ISERROR(value)**: Returns TRUE if value is an error, use in IF statements
Example: =IF(ISERROR(A1/B1),"Check Data",A1/B1)
- Other error functions: ISNA(), IFNA() (for #N/A specifically)

Would you like me to create practice examples or elaborate on any specific area?

Here are comprehensive formula-related Excel interview questions with detailed answers:
