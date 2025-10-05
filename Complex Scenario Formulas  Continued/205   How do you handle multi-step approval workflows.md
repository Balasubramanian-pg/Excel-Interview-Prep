### 205. **How do you handle multi-step approval workflows?**

**Current Approval Stage:**
=IFS(
Stage_1_Date="", "Pending Stage 1",
Stage_2_Date="", "Pending Stage 2",
Stage_3_Date="", "Pending Stage 3",
TRUE, "Approved"
)

**Days at Current Stage:**
=NETWORKDAYS(
MAX(Stage_1_Date, Stage_2_Date, Stage_3_Date),
TODAY()
)

**Total Approval Time:**
=NETWORKDAYS(Submission_Date, Final_Approval_Date)

**Approval Status Color:**
=IFS(
Status="Approved", "Green",
Days_Pending>SLA_Days, "Red",
Days_Pending>SLA_Days*0.8, "Yellow",
TRUE, "Green"
)

**Escalation Required:**
=AND(Status<>"Approved", Days_at_Stage>Escalation_Threshold)
