### 19. **How do you extract first/last name from full name?**

**First Name:** =LEFT(A1, FIND(" ", A1)-1)
**Last Name:** =RIGHT(A1, LEN(A1)-FIND(" ", A1))

For middle names, more complex:
=TRIM(MID(A1, FIND(" ", A1), FIND(" ", A1, FIND(" ", A1)+1)-FIND(" ", A1)))
