### 160. **Marketing: How do you calculate attribution models?**

**Last-Touch Attribution:**
All credit to last touchpoint: =IF(Touchpoint=Last_Touchpoint, Revenue, 0)

**First-Touch Attribution:**
All credit to first touchpoint: =IF(Touchpoint=First_Touchpoint, Revenue, 0)

**Linear Attribution:**
Equal credit across all touchpoints: =Revenue / Total_Touchpoints

**Time-Decay Attribution:**
More recent touchpoints get more credit:
=Revenue * (Power_Value^Days_Before_Conversion) / SUM(Power_Values)

**Position-Based (U-Shaped):**
40% first, 40% last, 20% distributed among middle:
=IFS(
Position=1, Revenue*0.4,
Position=Last, Revenue*0.4,
TRUE, Revenue*0.2/(Total_Touchpoints-2)
)
