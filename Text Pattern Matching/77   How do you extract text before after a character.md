### 77. **How do you extract text before/after a character?**

**Before character:**
=LEFT(A1, FIND("@", A1)-1)

**After character:**
=MID(A1, FIND("@", A1)+1, LEN(A1))

**Between two characters:**
=MID(A1, FIND("(", A1)+1, FIND(")", A1)-FIND("(", A1)-1)
