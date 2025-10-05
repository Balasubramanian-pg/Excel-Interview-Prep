### 111. **How do you perform hypothesis testing?**

**T-Test:**

- **T.TEST(array1, array2, tails, type):**
    - tails: 1 (one-tailed) or 2 (two-tailed)
    - type: 1 (paired), 2 (equal variance), 3 (unequal variance)

Example: =T.TEST(A1:A50, B1:B50, 2, 2)
Returns p-value for two-tailed test with equal variance

**Chi-Square Test:**

- **CHISQ.TEST(actual_range, expected_range):**
Example: =CHISQ.TEST(A1:A10, B1:B10)
Returns p-value
