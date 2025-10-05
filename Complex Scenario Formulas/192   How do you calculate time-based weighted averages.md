### 192. **How do you calculate time-based weighted averages?**

**Time-weighted average (for rates that change):**
=SUMPRODUCT(Values, Days_at_Value) / SUM(Days_at_Value)

**Volume-weighted average price (VWAP):**
=SUMPRODUCT(Price, Volume) / SUM(Volume)

**Exponential time decay:**
=SUMPRODUCT(Values, Decay_Factor^Days_Ago) / SUM(Decay_Factor^Days_Ago)
