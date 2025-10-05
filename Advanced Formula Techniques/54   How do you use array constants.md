### 54. **How do you use array constants?**

Create arrays directly in formulas using curly braces:

- Vertical: {1;2;3} (semicolons)
- Horizontal: {1,2,3} (commas)
- 2D: {1,2,3;4,5,6} (2 rows, 3 columns)

Examples:

- =SUM({1,2,3,4,5}) returns 15
- =VLOOKUP(A1, {"A","Apple";"B","Banana";"C","Cherry"}, 2, 0)
- =SUMPRODUCT((MONTH(A:A)={1,2,12})*(B:B)) sums B where A is Jan, Feb, or Dec
