### 210. **How do you handle currency hedging calculations?**

**Forward Rate:**
=Spot_Rate * (1 + Domestic_Rate*Days/360) / (1 + Foreign_Rate*Days/360)

**Hedge Effectiveness:**
=(Change_in_Hedge_Value / Change_in_Exposure_Value) * 100

**Natural Hedge Benefit:**
=ABS(Foreign_Receivables - Foreign_Payables) * Exchange_Rate_Volatility

**Hedge Ratio:**
=Value_of_Hedged_Position / Total_Foreign_Exposure

**Option Hedge Payoff:**
=MAX(0, Spot_Rate - Strike_Price) * Notional_Amount - Option_Premium
