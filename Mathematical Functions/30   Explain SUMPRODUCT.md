### 30. **Explain SUMPRODUCT**

Multiplies corresponding array elements and sums results:
Syntax: =SUMPRODUCT(array1, array2, ...)

Example: =SUMPRODUCT(A1:A10, B1:B10)
Multiplies A1*B1, A2*B2, etc., then sums all products

**Advanced uses:**

- Counting with criteria: =SUMPRODUCT((A1:A10="Yes")*(B1:B10>100))
- Weighted average: =SUMPRODUCT(scores, weights)/SUM(weights)
