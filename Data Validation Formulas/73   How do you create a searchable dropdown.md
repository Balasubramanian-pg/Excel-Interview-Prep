### 73. **How do you create a searchable dropdown?**

**Excel 365:**
Data Validation → List:
=FILTER(NamedRange, ISNUMBER(SEARCH(A1, NamedRange)))

As you type in A1, dropdown shows matching items.

**Older Excel:** Requires VBA or workarounds with helper columns
