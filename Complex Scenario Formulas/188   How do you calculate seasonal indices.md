### 188. **How do you calculate seasonal indices?**

**Classic Seasonal Index Method:**

1. Calculate moving average
2. Center moving average
3. Calculate ratio to moving average
4. Average ratios by period

**Formula for monthly index:**
=AVERAGE(IF(MONTH(Date_Range)=Month_Number, Actual/Moving_Avg))

**Seasonally adjusted value:**
=Actual_Value / Seasonal_Index

**Forecast with seasonality:**
=Trend_Value * Seasonal_Index
