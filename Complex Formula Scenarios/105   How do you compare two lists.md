### 105. **How do you compare two lists?**

**Items in List1 not in List2:**
=FILTER(List1, ISNA(XMATCH(List1, List2)))

**Items in both lists (intersection):**
=FILTER(List1, ISNUMBER(XMATCH(List1, List2)))

**All unique items (union):**
=UNIQUE(VSTACK(List1, List2))

These are the most comprehensive formula-related Excel questions you'll encounter in interviews! Would you like me to elaborate on any specific area or create practice examples?

Here are additional advanced formula topics and specialized scenarios:
