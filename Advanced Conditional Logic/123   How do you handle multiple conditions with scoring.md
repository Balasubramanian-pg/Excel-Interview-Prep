### 123. **How do you handle multiple conditions with scoring?**

**Weighted scoring system:**
=SUMPRODUCT(
(A1="Yes")*10,
(B1>100)*20,
(C1="Premium")*15,
(D1>=EOMONTH(TODAY(),-1))*5
)

Each TRUE condition adds its weight to total score
