### 196. **How do you calculate customer segmentation scores? (Continued)**

**Combined RFM Score:**
=CONCATENATE(Recency_Score, Frequency_Score, Monetary_Score)
Or: =R_Score & F_Score & M_Score

**RFM Segment Classification:**
=IFS(
RFM_Score="555", "Champions",
R_Score>=4*AND(F_Score>=4, M_Score>=4), "Loyal Customers",
R_Score>=4*AND(F_Score<=2, M_Score<=2), "Promising",
R_Score<=2*AND(F_Score>=4, M_Score>=4), "At Risk",
R_Score<=2*AND(F_Score<=2, M_Score>=4), "Can't Lose Them",
R_Score<=1, "Lost",
TRUE, "Need Attention"
)

**Weighted RFM Score:**
=(Recency_Score * 0.5) + (Frequency_Score * 0.3) + (Monetary_Score * 0.2)
