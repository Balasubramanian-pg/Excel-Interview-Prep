### 20. **How do you extract email domain?**

=MID(A1, FIND("@", A1)+1, LEN(A1))
Or: =RIGHT(A1, LEN(A1)-FIND("@", A1))
