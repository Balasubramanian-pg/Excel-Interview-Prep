### 177. **SaaS: How do you calculate MRR and ARR?**

**Monthly Recurring Revenue (MRR):**
=SUM(Active_Subscription_Values)

**Annual Recurring Revenue (ARR):**
=MRR * 12

**New MRR:**
=SUM(New_Subscriptions_This_Month)

**Expansion MRR:**
=SUM(Upsells + Cross_Sells)

**Contraction MRR:**
=SUM(Downgrades)

**Churned MRR:**
=SUM(Cancelled_Subscriptions)

**Net New MRR:**
=New_MRR + Expansion_MRR - Contraction_MRR - Churned_MRR
