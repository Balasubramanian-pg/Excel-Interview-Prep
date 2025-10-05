### 24. **What is DATEDIF and how do you use it?**

Calculates difference between two dates:
Syntax: =DATEDIF(start_date, end_date, unit)

Units:

- "Y": Complete years
- "M": Complete months
- "D": Days
- "YM": Months ignoring years
- "YD": Days ignoring years
- "MD": Days ignoring months and years

Example: =DATEDIF(A1, TODAY(), "Y") & " years, " & DATEDIF(A1, TODAY(), "YM") & " months"
