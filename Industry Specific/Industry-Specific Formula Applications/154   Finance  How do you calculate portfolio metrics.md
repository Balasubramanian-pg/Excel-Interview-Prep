### 154. **Finance: How do you calculate portfolio metrics?**

**Portfolio Return:**
=SUMPRODUCT(Weights, Returns)

**Portfolio Variance:**
=MMULT(MMULT(TRANSPOSE(Weights), Covariance_Matrix), Weights)

**Portfolio Standard Deviation:**
=SQRT(Portfolio_Variance)

**Sharpe Ratio:**
=(Portfolio_Return - Risk_Free_Rate) / Portfolio_StdDev

**Beta:**
=COVARIANCE.P(Stock_Returns, Market_Returns) / VAR.P(Market_Returns)

**Alpha (Jensen's):**
=Actual_Return - (Risk_Free_Rate + Beta*(Market_Return - Risk_Free_Rate))
