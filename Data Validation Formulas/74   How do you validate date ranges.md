### 74. **How do you validate date ranges?**

Data Validation → Custom:
=AND(A1>=TODAY(), A1<=TODAY()+30)

Only allows dates between today and 30 days from now.

**Business days only:**
=WEEKDAY(A1, 2)<=5

Rejects weekends (Saturday/Sunday)
