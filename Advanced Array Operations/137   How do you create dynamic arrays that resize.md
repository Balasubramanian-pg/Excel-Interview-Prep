### 137. **How do you create dynamic arrays that resize?**

**Spilling formula that expands:**
=FILTER(A:C, A:A<>"")

Automatically includes all non-empty rows

**With SEQUENCE for row numbers:**
=HSTACK(SEQUENCE(COUNTA(A:A)), FILTER(A:C, A:A<>""))

Adds row numbers that adjust automatically
