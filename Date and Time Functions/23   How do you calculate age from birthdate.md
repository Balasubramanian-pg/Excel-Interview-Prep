### 23. **How do you calculate age from birthdate?**

**Method 1:** =DATEDIF(birthdate, TODAY(), "Y")
**Method 2:** =INT((TODAY()-birthdate)/365.25)
**Method 3:** =YEARFRAC(birthdate, TODAY())

DATEDIF is hidden function but most accurate.
