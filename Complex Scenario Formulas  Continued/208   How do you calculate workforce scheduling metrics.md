### 208. **How do you calculate workforce scheduling metrics?**

**Coverage Ratio:**
=Scheduled_Staff / Required_Staff

**Schedule Efficiency:**
=(Productive_Hours / Total_Scheduled_Hours) * 100

**Overtime Hours:**
=MAX(0, Actual_Hours - Regular_Hours)

**Shift Premium:**
=IF(HOUR(Shift_Start)>=18, Hours*Rate*Shift_Premium, 0) +
IF(HOUR(Shift_Start)<6, Hours*Rate*Night_Premium, 0)

**Weekend Differential:**
=IF(WEEKDAY(Date,2)>=6, Hours*Rate*Weekend_Premium, 0)

**Consecutive Days Worked:**
=COUNTIF(OFFSET(Date_Column, -6, 0, 7, 1), "Scheduled")

**Fairness Index (schedule equity):**
=STDEV(Hours_by_Employee) / AVERAGE(Hours_by_Employee)
