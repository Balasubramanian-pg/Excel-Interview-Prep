### 180. **Technology: How do you calculate system performance?**

**Uptime Percentage:**
=(Total_Time - Downtime) / Total_Time * 100

**Availability (9s):**

- 99.9% = "three nines" = 8.76 hours downtime/year
- 99.99% = "four nines" = 52.56 minutes downtime/year

**Mean Time Between Failures (MTBF):**
=Total_Operating_Time / Number_of_Failures

**Mean Time To Repair (MTTR):**
=Total_Repair_Time / Number_of_Repairs

**Mean Time To Detect (MTTD):**
=Total_Detection_Time / Number_of_Incidents

**Error Rate:**
=(Errors / Total_Requests) * 100

**Response Time (Percentile):**
=PERCENTILE.INC(Response_Times, 0.95)
(95th percentile response time)
