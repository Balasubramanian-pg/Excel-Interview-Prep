### 18. **Explain all text manipulation functions**

**Joining Text:**

- **CONCATENATE(text1, text2, ...)**: Joins text (older method)
- **CONCAT(text1, text2, ...)**: Modern version, handles ranges
- **TEXTJOIN(delimiter, ignore_empty, text1, ...)**: Joins with delimiter
Example: =TEXTJOIN(", ", TRUE, A1:A5)
- **&** operator: =A1&" "&B1

**Extracting Text:**

- **LEFT(text, num_chars)**: Gets characters from left
Example: =LEFT(A1, 3) gets first 3 characters
- **RIGHT(text, num_chars)**: Gets characters from right
Example: =RIGHT(A1, 4) gets last 4 characters
- **MID(text, start_num, num_chars)**: Gets characters from middle
Example: =MID(A1, 5, 10) starts at position 5, takes 10 characters

**Finding Text:**

- **FIND(find_text, within_text, [start_num])**: Case-sensitive, returns position
Example: =FIND("@", A1) finds position of @
- **SEARCH(find_text, within_text, [start_num])**: Not case-sensitive, allows wildcards
Example: =SEARCH("excel", A1) finds "Excel", "EXCEL", etc.

**Modifying Text:**

- **UPPER(text)**: Converts to uppercase
- **LOWER(text)**: Converts to lowercase
- **PROPER(text)**: Capitalizes first letter of each word
- **TRIM(text)**: Removes extra spaces (leaves single spaces between words)
- **SUBSTITUTE(text, old_text, new_text, [instance])**: Replaces text
Example: =SUBSTITUTE(A1, "old", "new") replaces all "old" with "new"
- **REPLACE(old_text, start_num, num_chars, new_text)**: Replaces by position
Example: =REPLACE(A1, 1, 5, "New") replaces first 5 characters

**Other Text Functions:**

- **LEN(text)**: Returns length of text
- **REPT(text, number_times)**: Repeats text
Example: =REPT("*", 5) returns "*****"
- **TEXT(value, format_text)**: Formats numbers as text
Example: =TEXT(1234.5, "$#,##0.00") returns "$1,234.50"
- **VALUE(text)**: Converts text to number
