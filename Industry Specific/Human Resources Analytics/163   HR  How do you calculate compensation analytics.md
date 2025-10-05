### 163. **HR: How do you calculate compensation analytics?**

**Compa-Ratio:**
=Employee_Salary / Midpoint_of_Salary_Range

**Range Penetration:**
=(Employee_Salary - Range_Minimum) / (Range_Maximum - Range_Minimum)

**Pay Equity Analysis:**
=AVERAGE(IF(Gender="Female", Salary)) / AVERAGE(IF(Gender="Male", Salary))

**Compensation Increase Budget:**
=SUMIF(Employee_Status, "Active", Current_Salary) * Merit_Increase_Percentage

**Total Compensation:**
=Base_Salary + Bonus + Equity_Value + Benefits_Value
