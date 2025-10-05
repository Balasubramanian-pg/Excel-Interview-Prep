### 98. **How do you create a fiscal year calculator?**

If fiscal year starts in July:
=IF(MONTH(A1)>=7, YEAR(A1)+1, YEAR(A1))

**Fiscal quarter:**
=ROUNDUP((MONTH(A1)-6)/3, 0)
Adjust -6 based on fiscal year start

**Fiscal period (1-12):**
=IF(MONTH(A1)>=7, MONTH(A1)-6, MONTH(A1)+6)
