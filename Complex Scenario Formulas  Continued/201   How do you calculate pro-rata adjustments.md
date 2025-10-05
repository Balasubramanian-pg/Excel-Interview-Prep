### 201. **How do you calculate pro-rata adjustments?**

**Time-based pro-rata:**
=(Annual_Amount / 365) * Days_in_Period

**Percentage-based pro-rata:**
=Total_Amount * (Individual_Value / SUM(All_Values))

**Pro-rata refund:**
=Original_Amount * (Remaining_Days / Total_Contract_Days)

**Salary pro-rata (mid-month start):**
=(Annual_Salary / 12) * (EOMONTH(Start_Date,0) - Start_Date + 1) / DAY(EOMONTH(Start_Date,0))

**Partial period depreciation:**
=Annual_Depreciation * (Months_Owned / 12)
