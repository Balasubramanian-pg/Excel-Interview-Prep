### 164. **HR: How do you calculate headcount and FTE?**

**Full-Time Equivalent (FTE):**
=Hours_Worked / 40  (for weekly) or =Hours_Worked / 2080 (for annual)

**Total FTE:**
=SUM(FTE_Column)

**Average Headcount:**
=(SUM(Daily_Headcount) / Days_in_Period)

**Headcount Growth Rate:**
=(Current_Headcount - Previous_Headcount) / Previous_Headcount

**Span of Control:**
=COUNTIF(Manager_Column, Manager_Name)
