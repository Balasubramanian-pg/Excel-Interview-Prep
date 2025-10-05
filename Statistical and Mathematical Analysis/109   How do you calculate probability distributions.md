### 109. **How do you calculate probability distributions?**

**Normal Distribution:**

- **NORM.DIST(x, mean, standard_dev, cumulative):**
Example: =NORM.DIST(75, 70, 5, TRUE) returns probability ≤ 75
- **NORM.INV(probability, mean, standard_dev):**
Example: =NORM.INV(0.95, 70, 5) returns value at 95th percentile

**Binomial Distribution:**

- **BINOM.DIST(number_s, trials, probability_s, cumulative):**
Example: =BINOM.DIST(6, 10, 0.5, FALSE) probability of exactly 6 successes in 10 trials

**Poisson Distribution:**

- **POISSON.DIST(x, mean, cumulative):**
Example: =POISSON.DIST(5, 3, FALSE) probability of exactly 5 events when average is 3
