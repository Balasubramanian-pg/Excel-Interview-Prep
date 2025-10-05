### 203. **How do you create resource utilization tracking?**

**Utilization Rate:**
=(Billable_Hours / Total_Available_Hours) * 100

**Capacity Planning:**
=Total_Hours_Required / (Available_Resources * Hours_Per_Resource)

**Overbooking Calculation:**
=MAX(0, Scheduled_Hours - Available_Hours)

**Resource Efficiency:**
=(Actual_Output / Standard_Output) * 100

**Multi-resource allocation:**
=SUMIFS(Allocated_Hours, Resource, Current_Resource, Week, Current_Week) / Total_Hours_Available

**Forecast resource needs:**
=ROUNDUP(Projected_Hours / (Utilization_Target * Hours_Per_Person), 0)
