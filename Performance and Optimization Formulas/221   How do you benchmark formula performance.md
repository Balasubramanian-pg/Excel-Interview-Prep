### 221. **How do you benchmark formula performance?**

**Time calculation:**
Not directly in formulas, but measure recalc time:

- F9 to force recalculation
- Check Calculation tab in Options
- Use external timer for large datasets

**Array size check:**
=ROWS(Array) * COLUMNS(Array)

**Formula complexity indicator:**
=LEN(FORMULATEXT(A1))
(Longer formulas generally slower)

**Volatile function counter:**
=SUMPRODUCT(
--(ISNUMBER(SEARCH({"NOW","TODAY","RAND","OFFSET","INDIRECT"}, FORMULATEXT(A1))))
)
