### 121. **How do you create complex nested conditions?**

**Using SWITCH (cleaner than nested IFs):**
=SWITCH(A1,
"A", "Excellent",
"B", "Good",
"C", "Average",
"D", "Poor",
"F", "Fail",
"Invalid Grade")

**Multiple variable conditions:**
=SWITCH(TRUE,
AND(A1>90, B1="Y"), "Tier 1",
AND(A1>80, B1="Y"), "Tier 2",
AND(A1>70), "Tier 3",
"No Tier")
