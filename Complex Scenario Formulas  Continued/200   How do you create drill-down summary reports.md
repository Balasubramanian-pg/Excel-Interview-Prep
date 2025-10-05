### 200. **How do you create drill-down summary reports?**

**Conditional aggregation by level:**
=SUMIFS(Amount,
Category_Level_1, Selected_L1,
Category_Level_2, IF(Show_L2_Detail, Selected_L2, "*"),
Category_Level_3, IF(Show_L3_Detail, Selected_L3, "*")
)

**Dynamic row count:**
=COUNTA(UNIQUE(FILTER(Data, (L1=Selected_L1)*(L2=IF(Detail_Mode, Selected_L2, L2)))))

**Excel 365 - Expandable hierarchy:**
=IF(Expand_Flag,
FILTER(Detail_Data, Parent=Current_Item),
Current_Item
)
