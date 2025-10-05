### 122. **How do you create cascading conditions?**

**Priority-based logic:**
=IFS(
C1="Override", "Special",
A1>100, "High",
B1="Priority", "Medium",
TRUE, "Low"
)

First matching condition wins - order matters!
