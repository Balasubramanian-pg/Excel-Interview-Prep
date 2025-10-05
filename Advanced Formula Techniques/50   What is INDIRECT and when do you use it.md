### 50. **What is INDIRECT and when do you use it?**

Converts text string to cell reference:
Syntax: =INDIRECT(ref_text, [a1])

Examples:

- =INDIRECT("A" & ROW()) creates dynamic cell reference
- =INDIRECT(A1) where A1 contains "B5" returns value of B5
- =SUM(INDIRECT("Sheet" & A1 & "!A1:A10")) sums from different sheets

**Use cases:**

- Dynamic sheet references
- Creating cell references from text
- Building flexible formulas

**Warning:** INDIRECT is volatile (recalculates constantly), can slow workbooks
