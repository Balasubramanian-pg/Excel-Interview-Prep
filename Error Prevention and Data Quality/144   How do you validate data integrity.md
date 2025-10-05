### 144. **How do you validate data integrity?**

**Check for duplicates:**
=IF(COUNTIF($A$1:$A$1000,A1)>1, "Duplicate", "Unique")

**Verify referential integrity:**
=IF(ISNA(XMATCH(A1, MasterList)), "Missing in Master", "OK")

**Identify orphaned records:**
=FILTER(ChildTable, ISNA(XMATCH(ChildID, ParentID)))

**Check for missing sequence numbers:**
=FILTER(SEQUENCE(MAX(A:A)), ISNA(XMATCH(SEQUENCE(MAX(A:A)), A:A)))
