### 39. **Explain FILTER function**

Returns array filtered by criteria:
Syntax: =FILTER(array, include, [if_empty])

Example: =FILTER(A1:C100, B1:B100>1000, "No results")
Returns all rows where column B is greater than 1000

Multiple criteria with AND:
=FILTER(A1:C100, (B1:B100>1000)*(C1:C100="Active"))

Multiple criteria with OR:
=FILTER(A1:C100, (B1:B100>1000)+(C1:C100="VIP"))
