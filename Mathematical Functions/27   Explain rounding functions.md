### 27. **Explain rounding functions**

- **ROUND(number, num_digits)**: Rounds to specified decimals
Example: =ROUND(3.456, 2) returns 3.46
- **ROUNDUP(number, num_digits)**: Always rounds up
Example: =ROUNDUP(3.451, 2) returns 3.46
- **ROUNDDOWN(number, num_digits)**: Always rounds down
Example: =ROUNDDOWN(3.459, 2) returns 3.45
- **MROUND(number, multiple)**: Rounds to nearest multiple
Example: =MROUND(23, 5) returns 25
- **CEILING(number, significance)**: Rounds up to multiple
Example: =CEILING(23, 5) returns 25
- **FLOOR(number, significance)**: Rounds down to multiple
Example: =FLOOR(23, 5) returns 20
- **INT(number)**: Rounds down to nearest integer
- **TRUNC(number, [num_digits])**: Truncates (removes decimals)
