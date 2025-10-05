### 90. **How do you use SEQUENCE for dynamic ranges?**

**Create row numbers:**
=SEQUENCE(10) generates 1-10

**Create date series:**
=TODAY() + SEQUENCE(7) - 1
Generates next 7 days

**Create multiplication table:**
=SEQUENCE(10) * SEQUENCE(1, 10)
Generates 10x10 multiplication table

**Dynamic month list:**
=TEXT(DATE(2025, SEQUENCE(12), 1), "MMM")
Generates Jan, Feb, Mar... Dec
