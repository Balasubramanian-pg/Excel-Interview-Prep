### 46. **When should you use IFERROR vs IFNA?**

- **IFNA**: Use for lookup functions (VLOOKUP, XLOOKUP, MATCH) where #N/A is expected when item not found
- **IFERROR**: Use for calculations where multiple error types possible

IFNA is more precise - it won't hide formula errors like #REF! or #VALUE!
