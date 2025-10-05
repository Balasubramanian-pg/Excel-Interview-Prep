### 185. **BI: How do you calculate market basket analysis?**

**Support (Item Frequency):**
=COUNTIF(Transaction_Items, Item) / Total_Transactions

**Confidence (A → B):**
=COUNTIFS(Trans_Has_A, TRUE, Trans_Has_B, TRUE) / COUNTIF(Trans_Has_A, TRUE)

**Lift (A & B together):**
=(Support_AB) / (Support_A * Support_B)

**Interpretation:**

- Lift > 1: Items purchased together more than expected
- Lift = 1: No relationship
- Lift < 1: Negative correlation
