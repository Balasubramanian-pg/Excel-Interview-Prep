### 115. **How do you split text into columns with formulas?**

**Excel 365 - TEXTSPLIT:**
=TEXTSPLIT(A1, ",")
Splits by comma into columns

**With both row and column delimiters:**
=TEXTSPLIT(A1, ",", ";")
Comma separates columns, semicolon separates rows

**Older Excel:**

- First item: =LEFT(A1, FIND(",", A1)-1)
- Second item: =MID(A1, FIND(",", A1)+1, FIND(",", A1, FIND(",", A1)+1)-FIND(",", A1)-1)
- Last item: =RIGHT(A1, LEN(A1)-FIND("~", SUBSTITUTE(A1, ",", "~", LEN(A1)-LEN(SUBSTITUTE(A1, ",", "")))))
