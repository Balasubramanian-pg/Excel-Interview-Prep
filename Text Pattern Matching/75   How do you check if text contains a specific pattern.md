### 75. **How do you check if text contains a specific pattern?**

**Contains any text:**
=ISNUMBER(SEARCH("text", A1))

**Starts with specific text:**
=LEFT(A1, LEN("text"))="text"

**Ends with specific text:**
=RIGHT(A1, LEN("text"))="text"

**Matches pattern (wildcards):**
=COUNTIF(A1, "*pattern*")>0
