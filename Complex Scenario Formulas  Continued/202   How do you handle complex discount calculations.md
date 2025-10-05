### 202. **How do you handle complex discount calculations?**

**Volume-based tiered discount:**
=SUMPRODUCT(
--(Quantity >= Tier_Minimums),
MIN(Quantity, Tier_Maximums) - Tier_Minimums + 1,
Base_Price * (1 - Tier_Discounts)
)

**Cumulative discount (discount on discount):**
=Base_Price * (1 - Discount1) * (1 - Discount2) * (1 - Discount3)

**Best discount selector:**
=Base_Price * (1 - MAX(Volume_Discount, Loyalty_Discount, Promotional_Discount))

**Bundle discount:**
=IF(Has_Product_A * Has_Product_B,
(Price_A + Price_B) * (1 - Bundle_Discount),
Price_A * Has_Product_A + Price_B * Has_Product_B
)

**Early payment discount:**
=IF(Payment_Date <= Invoice_Date + Early_Pay_Days,
Invoice_Amount * (1 - Early_Pay_Discount),
Invoice_Amount
)
