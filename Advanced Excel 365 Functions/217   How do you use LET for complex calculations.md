### 217. **How do you use LET for complex calculations?**

**Avoid recalculation:**
=LET(
raw_data, A1:A100,
cleaned, FILTER(raw_data, raw_data<>""),
mean, AVERAGE(cleaned),
stdev, STDEV(cleaned),
z_scores, (cleaned - mean) / stdev,
FILTER(cleaned, ABS(z_scores)<3)
)

**Multi-step business calculation:**
=LET(
revenue, A1,
cogs, B1,
opex, C1,
gross_profit, revenue - cogs,
gross_margin, gross_profit / revenue,
operating_profit, gross_profit - opex,
operating_margin, operating_profit / revenue,
HSTACK(gross_profit, gross_margin, operating_profit, operating_margin)
)

**Nested calculations:**
=LET(
x, A1,
y, B1,
sum_xy, x + y,
product_xy, x * y,
ratio, x / y,
final, (sum_xy * product_xy) / ratio,
final
)
