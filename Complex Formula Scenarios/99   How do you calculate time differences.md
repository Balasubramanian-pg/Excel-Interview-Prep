### 99. **How do you calculate time differences?**

**Simple time difference:**
=B1-A1 (format as time)

**Hours between times:**
=(B1-A1)*24

**Minutes between times:**
=(B1-A1)*1440

**Across midnight:**
=IF(B1<A1, 1+B1-A1, B1-A1)

**Business hours only (9 AM - 5 PM):**
Complex formula considering start/end times, lunch breaks, etc.
