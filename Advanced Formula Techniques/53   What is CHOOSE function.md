### 53. **What is CHOOSE function?**

Returns value from list based on index:
Syntax: =CHOOSE(index_num, value1, value2, ...)

Example: =CHOOSE(2, "Red", "Blue", "Green") returns "Blue"

**Use cases:**

- Convert numbers to text: =CHOOSE(MONTH(A1), "Jan", "Feb", "Mar", ...)
- Dynamic calculations: =CHOOSE(A1, B1+C1, B1*C1, B1/C1)
- In combination with MATCH for advanced lookups
