### 155. **Finance: How do you calculate depreciation?**

**Straight-Line:**
=SLN(cost, salvage, life)
Example: =SLN(100000, 10000, 10)

**Declining Balance:**
=DB(cost, salvage, life, period, [month])
Example: =DB(100000, 10000, 10, 1)

**Double-Declining Balance:**
=DDB(cost, salvage, life, period, [factor])
Example: =DDB(100000, 10000, 10, 1, 2)

**Sum-of-Years' Digits:**
=SYD(cost, salvage, life, period)
Example: =SYD(100000, 10000, 10, 1)

**Variable Declining Balance:**
=VDB(cost, salvage, life, start_period, end_period, [factor], [no_switch])
