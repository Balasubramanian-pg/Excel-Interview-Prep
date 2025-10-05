### 135. **How do you create aging buckets for receivables?**

=IFS(
TODAY()-A1<=30, "Current",
TODAY()-A1<=60, "31-60 Days",
TODAY()-A1<=90, "61-90 Days",
TODAY()-A1<=120, "91-120 Days",
TRUE, "Over 120 Days"
)

**For aging summary:**
=SUMIFS(Amount, InvoiceDate, ">="&TODAY()-30, InvoiceDate, "<"&TODAY())
