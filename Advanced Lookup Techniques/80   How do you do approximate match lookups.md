### 80. **How do you do approximate match lookups?**

VLOOKUP with TRUE (or 1) as 4th argument:
=VLOOKUP(A1, Table, 2, TRUE)

**Requirements:**

- Lookup column must be sorted ascending
- Returns largest value less than or equal to lookup value

**Use case:** Grade ranges, tax brackets, commission tiers
