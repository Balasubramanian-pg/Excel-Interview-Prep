### 141. **How do you create data validation with complex rules?**

**Prevent overlapping date ranges:**
=COUNTIFS(StartDates, "<="&B1, EndDates, ">="&A1)=0

Apply to start/end date pair

**Ensure sum equals specific value:**
=SUM($A$1:$A$10)=100

**Validate email format:**
=AND(LEN(A1)>0, ISNUMBER(FIND("@",A1)), ISNUMBER(FIND(".",A1)), FIND("@",A1)<FIND(".",A1))

**Prevent weekends:**
=AND(WEEKDAY(A1,2)<=5, A1>=TODAY())
