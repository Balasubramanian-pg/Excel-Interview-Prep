### 117. **How do you create custom sorting orders?**

Use SORTBY with XMATCH and custom order list:

Custom order: {"High", "Medium", "Low"}
=SORTBY(A1:B100, XMATCH(B1:B100, {"High","Medium","Low"}))

**For multiple columns with custom orders:**
=SORTBY(A1:C100,
XMATCH(B1:B100, CustomOrder1), 1,
XMATCH(C1:C100, CustomOrder2), 1)
