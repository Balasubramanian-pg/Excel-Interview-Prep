### 63. **How do you sum every nth row?**

=SUMPRODUCT((MOD(ROW(A1:A100)-ROW(A1), n)=0)*(A1:A100))

Where n is the interval (3 for every 3rd row)

**Specific example - every 3rd row:**
=SUMPRODUCT((MOD(ROW(A1:A100), 3)=0)*(A1:A100))
