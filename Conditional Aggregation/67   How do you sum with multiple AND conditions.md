### 67. **How do you sum with multiple AND conditions?**

Use SUMIFS:
=SUMIFS(D:D, A:A, "West", B:B, ">1000", C:C, "Active")

Sums column D where:

- Column A = "West" AND
- Column B > 1000 AND
- Column C = "Active"
