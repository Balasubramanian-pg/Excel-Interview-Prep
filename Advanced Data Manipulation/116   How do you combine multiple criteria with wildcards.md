### 116. **How do you combine multiple criteria with wildcards?**

**Multiple wildcards in COUNTIFS:**
=COUNTIFS(A:A, "Apple*", B:B, "*Red*")
Counts where A starts with "Apple" AND B contains "Red"

**Complex pattern matching:**
=SUMPRODUCT((ISNUMBER(SEARCH("keyword1", A:A)) + ISNUMBER(SEARCH("keyword2", A:A)) > 0) * B:B)
Sums B where A contains keyword1 OR keyword2
