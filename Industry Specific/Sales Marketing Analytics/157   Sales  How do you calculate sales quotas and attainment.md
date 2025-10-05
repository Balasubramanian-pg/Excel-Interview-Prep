### 157. **Sales: How do you calculate sales quotas and attainment?**

**Quota Attainment:**
=Actual_Sales / Quota

**Commission Calculation (Tiered):**
=IFS(
Attainment<0.5, Actual_Sales*0.02,
Attainment<0.8, Actual_Sales*0.05,
Attainment<1.0, Actual_Sales*0.08,
Attainment<1.2, Actual_Sales*0.10,
TRUE, Actual_Sales*0.12
)

**Accelerated Commission:**
=IF(Attainment>1,
Quota*Base_Rate + (Actual_Sales-Quota)*Accelerated_Rate,
Actual_Sales*Base_Rate
)

**Quarter-to-Date Attainment:**
=SUMIFS(Sales, Date, ">="&DATE(YEAR(TODAY()), QUARTER(TODAY())*3-2, 1), Date, "<="&TODAY()) / Quarterly_Quota
