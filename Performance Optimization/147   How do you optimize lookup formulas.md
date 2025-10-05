### 147. **How do you optimize lookup formulas?**

**Instead of VLOOKUP:**
Use INDEX-MATCH (faster on large datasets)

**Instead of multiple nested IFs:**
Use SWITCH or IFS

**Instead of SUMIF with entire columns:**
Use specific ranges: =SUMIF(A1:A1000, criteria, B1:B1000)

**Use tables:** Structured references are more efficient
