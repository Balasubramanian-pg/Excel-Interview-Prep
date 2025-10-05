### 152. **Finance: How do you calculate bond pricing and yields?**

**Bond Price:**
=PV(yield/2, years*2, -coupon/2, -face_value)
(Dividing by 2 for semi-annual payments)

**Current Yield:**
=Annual_Coupon_Payment / Current_Market_Price

**Yield to Maturity (YTM):**
=YIELD(settlement, maturity, rate, pr, redemption, frequency, [basis])

**Duration (Macaulay):**
=DURATION(settlement, maturity, coupon, yld, frequency, [basis])

**Modified Duration:**
=MDURATION(settlement, maturity, coupon, yld, frequency, [basis])
