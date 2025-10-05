### 76. **How do you extract numbers from text?**

**Excel 365 (array formula):**
=SUMPRODUCT(MID(0&A1, LARGE(INDEX(ISNUMBER(--MID(A1, ROW($1:$99), 1)) * ROW($1:$99), 0), ROW($1:$99))+1, 1) * 10^ROW($1:$99)/10)

**Simpler for consistent formats:**
If "ABC123" → =VALUE(RIGHT(A1, 3))

**Best practice:** Use Power Query or VBA for complex extractions
