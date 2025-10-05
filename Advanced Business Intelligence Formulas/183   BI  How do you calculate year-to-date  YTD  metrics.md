### 183. **BI: How do you calculate year-to-date (YTD) metrics?**

**YTD Sum:**
=SUMIFS(Sales, Date, ">="&DATE(YEAR(TODAY()),1,1), Date, "<="&TODAY())

**YTD Average:**
=AVERAGEIFS(Sales, Date, ">="&DATE(YEAR(TODAY()),1,1), Date, "<="&TODAY())

**YTD vs Prior YTD:**
=(Current_YTD - Prior_YTD) / Prior_YTD

**YTD with Fiscal Year:**
=SUMIFS(Sales, Date, ">="&Fiscal_Year_Start, Date, "<="&TODAY())

**Dynamic YTD (Excel 365):**
=SUM(FILTER(Sales, (YEAR(Dates)=YEAR(TODAY()))*(Dates<=TODAY())))
