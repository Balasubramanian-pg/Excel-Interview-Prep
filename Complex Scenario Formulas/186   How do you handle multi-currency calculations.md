### 186. **How do you handle multi-currency calculations?**

**Convert to Base Currency:**
=Amount * XLOOKUP(Currency, Currency_Table, Exchange_Rate)

**Multi-step conversion:**
=Amount / Source_Rate * Target_Rate

**With historical rates:**
=Amount * XLOOKUP(1, (Currency_Table_Currency=Currency)*(Currency_Table_Date<=Trans_Date), Currency_Table_Rate)

**Average exchange rate for period:**
=AVERAGEIFS(Rates, Currency_Col, "EUR", Date_Col, ">="&Start, Date_Col, "<="&End)
