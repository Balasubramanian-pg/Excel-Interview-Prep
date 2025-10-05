### 94. **What is VSTACK and HSTACK?**

Stack arrays vertically or horizontally:

**VSTACK(array1, array2, ...):**
=VSTACK(A1:C5, A10:C15)
Stacks two ranges on top of each other

**HSTACK(array1, array2, ...):**
=HSTACK(A1:A5, C1:C5)
Places ranges side by side

**Combine with other functions:**
=VSTACK("Headers", FILTER(Data, Criteria))
