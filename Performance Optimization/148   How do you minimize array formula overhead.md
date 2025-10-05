### 148. **How do you minimize array formula overhead?**

**Prefer dynamic arrays (Excel 365) over CSE arrays**

**Use SUMPRODUCT instead of SUM(IF()):**
Better: =SUMPRODUCT((A:A="X")*(B:B))
Avoid: =SUM(IF(A:A="X",B:B))

**Limit array sizes:**
Specify exact ranges instead of entire columns when possible
