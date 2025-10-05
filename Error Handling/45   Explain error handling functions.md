### 45. **Explain error handling functions**

- **IFERROR(value, value_if_error)**: Catches all errors
Example: =IFERROR(A1/B1, 0) returns 0 if division errors
- **IFNA(value, value_if_na)**: Catches only #N/A
Example: =IFNA(VLOOKUP(A1, D:E, 2, 0), "Not Found")
- **ISERROR(value)**: Returns TRUE if any error
- **ISNA(value)**: Returns TRUE if #N/A
- **ISERR(value)**: Returns TRUE if any error except #N/A

**Best practice:** Use IFNA for lookups, IFERROR for calculations
