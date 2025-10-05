### 207. **How do you create allocation matrices?**

**Cost allocation by driver:**
=Total_Cost_Pool * (Department_Driver / SUM(All_Drivers))

**Step-down allocation:**

```
Dept_A_Allocation = Direct_Cost_A
Dept_B_Allocation = Direct_Cost_B + (Dept_A_Allocation * B_Uses_A%)
Dept_C_Allocation = Direct_Cost_C + (Dept_A_Allocation * C_Uses_A%) + (Dept_B_Allocation * C_Uses_B%)

```

**Matrix allocation (simultaneous):**
Requires solving system of equations: =MMULT(MINVERSE(Allocation_Matrix), Direct_Costs)

**Activity-based costing:**
=SUMPRODUCT(Activity_Costs, Activity_Drivers) / Total_Units
