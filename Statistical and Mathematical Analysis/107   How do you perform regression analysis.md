### 107. **How do you perform regression analysis?**

**Simple linear regression (slope and intercept):**

- **SLOPE(known_y's, known_x's):** Returns slope (m in y=mx+b)
Example: =SLOPE(B1:B100, A1:A100)
- **INTERCEPT(known_y's, known_x's):** Returns y-intercept (b)
Example: =INTERCEPT(B1:B100, A1:A100)

**Predict values:**
=SLOPE(B:B, A:A) * NewX + INTERCEPT(B:B, A:A)

**R-squared (goodness of fit):**
=RSQ(known_y's, known_x's)
Returns value 0-1 (closer to 1 = better fit)

**FORECAST.LINEAR(x, known_y's, known_x's):**
Predicts y value for given x
Example: =FORECAST.LINEAR(15, B1:B100, A1:A100)
