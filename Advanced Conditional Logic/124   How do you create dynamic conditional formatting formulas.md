### 124. **How do you create dynamic conditional formatting formulas?**

**Highlight entire row based on cell value:**
=$E1="Complete"
Apply to $A$1:$Z$1000

**Alternate row shading:**
=MOD(ROW(),2)=0

**Highlight duplicates in column:**
=COUNTIF($A$1:$A1,$A1)>1

**Highlight dates within next 7 days:**
=AND(A1>=TODAY(), A1<=TODAY()+7)

**Highlight top 10% of values:**
=A1>=PERCENTILE($A$1:$A$100,0.9)
