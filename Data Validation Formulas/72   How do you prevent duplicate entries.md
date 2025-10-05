### 72. **How do you prevent duplicate entries?**

Data Validation → Custom:
=COUNTIF($A$1:$A$1000, A1)=1

Apply to range A1:A1000. This prevents entering a value that already exists.

**Allow first entry, prevent subsequent:**
=COUNTIF($A$1:A1, A1)=1
