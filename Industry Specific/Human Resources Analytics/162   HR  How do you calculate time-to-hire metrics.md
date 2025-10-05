### 162. **HR: How do you calculate time-to-hire metrics?**

**Time to Fill:**
=Filled_Date - Requisition_Open_Date

**Time to Hire:**
=Hired_Date - Application_Date

**Average Time to Fill by Department:**
=AVERAGEIF(Department_Range, "Engineering", Time_to_Fill_Range)

**Offer Acceptance Rate:**
=Offers_Accepted / Offers_Extended

**Source of Hire Effectiveness:**
=COUNTIF(Source_Range, "LinkedIn") / COUNTIF(Status_Range, "Hired")
