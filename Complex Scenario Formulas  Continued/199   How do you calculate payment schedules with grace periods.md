### 199. **How do you calculate payment schedules with grace periods?**

**Due Date with Grace Period:**
=WORKDAY(Invoice_Date, Payment_Terms, Holidays) + Grace_Days

**Late Fee Calculation:**
=IF(Payment_Date > Grace_Date,
Invoice_Amount * Late_Fee_Rate * NETWORKDAYS(Grace_Date, Payment_Date) / 365,
0
)

**Payment Status:**
=IFS(
Payment_Date="", IF(TODAY()>Grace_Date, "Overdue", "Pending"),
Payment_Date<=Due_Date, "On Time",
Payment_Date<=Grace_Date, "Within Grace",
TRUE, "Late"
)

**Days Past Due:**
=MAX(0, NETWORKDAYS(Grace_Date, IF(Payment_Date="", TODAY(), Payment_Date)))
