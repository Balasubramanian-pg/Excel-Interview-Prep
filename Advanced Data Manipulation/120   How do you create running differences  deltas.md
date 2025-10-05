### 120. **How do you create running differences (deltas)?**

**Simple difference:**
=A2-A1
Drag down from second row

**Percentage change:**
=(A2-A1)/A1
Format as percentage

**Year-over-year change:**
=A13-A1  (if monthly data, row 13 is same month previous year)

**Excel 365 - for entire column:**
=DROP(A:A, 1) - DROP(A:A, -1, 1)
Returns differences between consecutive values
