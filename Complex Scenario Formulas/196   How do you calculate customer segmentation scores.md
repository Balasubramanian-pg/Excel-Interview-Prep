### 196. **How do you calculate customer segmentation scores?**

**RFM Score:**

- Recency Score: =IFS(Days_Since_Purchase<=30, 5, Days<=90, 4, Days<=180, 3, Days<=365, 2, TRUE, 1)
- Frequency Score: =IFS(Purchase_Count>=10, 5, >=7, 4, >=4, 3, >=2, 2, TRUE, 1)
- Monetary Score: =IFS(Total_Spent>=10000, 5, >=5000, 4, >=2000, 3, >=500, 2, TRUE, 1
