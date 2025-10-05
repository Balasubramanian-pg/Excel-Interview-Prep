### 187. **How do you create dynamic budget allocation?**

**Pro-rata allocation:**
=Total_Budget * (Department_Employees / Total_Employees)

**Weighted allocation:**
=Total_Budget * (Department_Revenue / Total_Revenue) * Weight_Factor

**Tiered allocation:**
=IFS(
Revenue<1000000, Base_Budget,
Revenue<5000000, Base_Budget + (Revenue-1000000)*0.10,
TRUE, Base_Budget + 400000 + (Revenue-5000000)*0.05
)
