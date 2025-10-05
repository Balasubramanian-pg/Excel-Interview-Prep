### 206. **How do you calculate tax in multi-jurisdictional scenarios?**

**Compound Tax (Federal + State):**
=Amount * (1 + Federal_Rate) * (1 + State_Rate) - Amount

**Alternative (if state tax is on subtotal):**
=Amount * (Federal_Rate + State_Rate + Federal_Rate*State_Rate)

**Cascading Tax:**

- Federal: =Amount * Federal_Rate
- State on Federal: =(Amount + Federal_Tax) * State_Rate
- Total: =Federal_Tax + State_Tax

**Location-based tax lookup:**
=Amount * XLOOKUP(Zip_Code, Tax_Table_Zip, Tax_Table_Rate, 0)

**Tax exclusive to inclusive:**
=Price * (1 + Tax_Rate)

**Tax inclusive to exclusive:**
=Price / (1 + Tax_Rate)
