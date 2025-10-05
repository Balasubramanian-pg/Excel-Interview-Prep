### 204. **How do you calculate service level agreements (SLA)?**

**SLA Compliance Percentage:**
=(Tickets_Within_SLA / Total_Tickets) * 100

**Time Remaining on SLA:**
=SLA_Deadline - NOW()

**SLA Breach Warning:**
=IF((SLA_Deadline - NOW())*24 <= Warning_Hours, "WARNING", "OK")

**Weighted SLA (by priority):**
=SUMPRODUCT(
(Priority_Range={"Critical","High","Medium","Low"}),
(Within_SLA_Range),
{0.4, 0.3, 0.2, 0.1}
) / SUMPRODUCT((Priority_Range={"Critical","High","Medium","Low"}), {0.4, 0.3, 0.2, 0.1})

**Business Hours SLA:**
=NETWORKDAYS.INTL(Start_Time, End_Time, 1, Holidays) * 8 -
HOUR(Start_Time) + HOUR(End_Time)
