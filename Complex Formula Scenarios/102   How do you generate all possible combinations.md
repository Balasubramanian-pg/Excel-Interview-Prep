### 102. **How do you generate all possible combinations?**

**Excel 365:** Use nested SEQUENCE:
=SEQUENCE(n) & "-" & SEQUENCE(1, m)

For full Cartesian product, more complex LAMBDA required

**Older Excel:** Requires VBA or manual helper columns
