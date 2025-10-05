### 119. **How do you group data into bins/buckets?**

**Using FREQUENCY (array formula):**
=FREQUENCY(Data, Bins)

Example: =FREQUENCY(A1:A100, {50;100;150;200})
Counts values in ranges: 0-50, 51-100, 101-150, 151-200, >200

**Using IFS for categorization:**
=IFS(A1<=50, "Low", A1<=100, "Medium", A1<=150, "High", TRUE, "Very High")

**Excel 365 with SWITCH:**
=SWITCH(TRUE, A1<=50, "Low", A1<=100, "Medium", A1<=150, "High", "Very High")
