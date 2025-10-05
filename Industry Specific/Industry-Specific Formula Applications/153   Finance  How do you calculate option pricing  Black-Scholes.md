### 153. **Finance: How do you calculate option pricing (Black-Scholes)?**

**Components needed:**

- S = Current stock price
- K = Strike price
- T = Time to expiration (years)
- r = Risk-free rate
- σ = Volatility

**d1 formula:**
=(LN(S/K) + (r + σ^2/2)*T) / (σ*SQRT(T))

**d2 formula:**
=d1 - σ*SQRT(T)

**Call Option Price:**
=S*NORM.S.DIST(d1, TRUE) - K*EXP(-r*T)*NORM.S.DIST(d2, TRUE)

**Put Option Price:**
=K*EXP(-r*T)*NORM.S.DIST(-d2, TRUE) - S*NORM.S.DIST(-d1, TRUE)
