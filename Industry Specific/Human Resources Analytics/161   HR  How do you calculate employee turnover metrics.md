### 161. **HR: How do you calculate employee turnover metrics?**

**Turnover Rate:**
=(Number_of_Separations / Average_Number_of_Employees) * 100

**Average Employees:**
=(Beginning_Headcount + Ending_Headcount) / 2

**Voluntary vs Involuntary Turnover:**
=COUNTIFS(Separation_Type, "Voluntary", Separation_Date, ">="&Start, Separation_Date, "<="&End) / Avg_Employees

**90-Day Turnover (New Hire):**
=COUNTIFS(Hire_Date, ">="&Start, Separation_Date, "<="&Hire_Date+90) / COUNTIFS(Hire_Date, ">="&Start)

**Annualized Turnover:**
=(Monthly_Separations * 12) / Average_Headcount

**Retention Rate:**
=1 - Turnover_Rate
