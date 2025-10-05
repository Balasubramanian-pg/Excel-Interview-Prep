### 7. **Explain all logical functions**

- **IF(test, true_value, false_value)**: Basic condition
Example: =IF(A1>100, "High", "Low")
- **AND(logical1, logical2, ...)**: Returns TRUE if all conditions are TRUE
Example: =IF(AND(A1>50, B1<100), "Valid", "Invalid")
- **OR(logical1, logical2, ...)**: Returns TRUE if any condition is TRUE
Example: =IF(OR(A1="Yes", B1="Yes"), "Approved", "Denied")
- **NOT(logical)**: Reverses TRUE/FALSE
Example: =IF(NOT(A1=""), "Has Value", "Empty")
- **XOR(logical1, logical2, ...)**: Returns TRUE if odd number of conditions are TRUE
Example: =XOR(A1>50, B1>50) true if only one is greater than 50
- **IFS(test1, value1, test2, value2, ...)**: Multiple conditions without nesting (Excel 2016+)
Example: =IFS(A1>=90, "A", A1>=80, "B", A1>=70, "C", TRUE, "F")
