### 167. **Operations: How do you calculate delivery performance?**

**On-Time Delivery (OTD):**
=COUNTIF(Delivery_Status, "On Time") / COUNT(Total_Deliveries)

**On-Time In-Full (OTIF):**
=COUNTIFS(On_Time, "Yes", In_Full, "Yes") / Total_Orders

**Perfect Order Rate:**
=COUNTIFS(On_Time, "Yes", Complete, "Yes", Damage_Free, "Yes", Doc_Accurate, "Yes") / Total_Orders

**Fill Rate:**
=Units_Delivered / Units_Ordered

**Backorder Rate:**
=Units_on_Backorder / Total_Units_Ordered
