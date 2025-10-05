### 146. **What formulas are volatile and should be minimized?**

**Volatile functions (recalculate every change):**

- NOW(), TODAY()
- RAND(), RANDBETWEEN()
- OFFSET()
- INDIRECT()
- INFO()

**Best practices:**

- Replace OFFSET with INDEX where possible
- Replace INDIRECT with direct references
- Calculate NOW() once in a cell and reference that cell
- Use RANDARRAY in Excel 365 instead of RAND in many cells
