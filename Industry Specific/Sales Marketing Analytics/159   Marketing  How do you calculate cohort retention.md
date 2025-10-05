### 159. **Marketing: How do you calculate cohort retention?**

**Month-over-Month Retention:**
=COUNTIFS(Customer_ID, Month_0_IDs, Active_Month, Month_N) / COUNT(Month_0_IDs)

**Cohort Analysis Formula:**
Structure with cohort month in rows, months since acquisition in columns:
=COUNTIFS(Cohort_Range, $A2, Month_Range, B$1) / COUNTIF(Cohort_Range, $A2)

**Cumulative Retention:**
=SUMIFS(Still_Active, Cohort, $A2, Month, "<="&B$1) / COUNTIF(Cohort, $A2)
