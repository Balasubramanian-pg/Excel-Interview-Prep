### 136. **How do you create cross-tabulation (pivot-like) formulas?**

**Sum by two criteria (manual pivot):**
=SUMIFS($D:$D, $A:$A, $G2, $B:$B, H$1)

Where G2 is row header, H1 is column header

**Excel 365 with GROUPBY (if available):**
This would require the PIVOT function when available

**Current Excel 365 workaround:**
Use combination of UNIQUE and SUMIFS
