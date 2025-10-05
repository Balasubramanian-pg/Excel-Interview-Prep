### 189. **How do you create ABC analysis (Pareto)?**

**Cumulative percentage:**
=(SUM($B$2:B2)/SUM($B$2:$B$1000))

**ABC Classification:**
=IFS(
Cumulative_Pct<=0.80, "A",
Cumulative_Pct<=0.95, "B",
TRUE, "C"
)

**Sort first by value descending, then apply cumulative formula**
