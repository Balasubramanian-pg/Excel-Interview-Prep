### 145. **How do you create checksums or hash validations?**

**Simple checksum (sum of digits):**
=SUMPRODUCT(--MID(A1,ROW(INDIRECT("1:"&LEN(A1))),1))

**Modulo-based check digit:**
=MOD(SUMPRODUCT(--MID(A1,ROW(INDIRECT("1:"&LEN(A1))),1)*{1,3}),10)

**Row-level validation:**
=IF(SUM(B1:F1)=G1, "Valid", "Error")
