### 181. **BI: How do you create rolling/moving averages?**

**Simple Moving Average (SMA):**
=AVERAGE(OFFSET(A1, COUNT($A$1:A1)-Period, 0, Period, 1))

**Weighted Moving Average:**
=SUMPRODUCT(OFFSET(A1, COUNT($A$1:A1)-Period, 0, Period, 1), Weights) / SUM(Weights)

**Exponential Moving Average (EMA):**
=IF(ROW()=2, A2, A2*Smoothing + EMA_Previous*(1-Smoothing))
Where Smoothing = 2/(Period+1)

**Excel 365 Dynamic:**
=BYROW(SEQUENCE(ROWS(Data)-Period+1), LAMBDA(r, AVERAGE(INDEX(Data, r):INDEX(Data, r+Period-1))))
