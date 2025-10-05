### 133. **How do you calculate compound annual growth rate (CAGR)?**

**Standard formula:**
=((Ending_Value / Beginning_Value)^(1/Number_of_Years)) - 1

**Using RRI function:**
=RRI(years, -beginning_value, ending_value)

**Example:**
=((B10/B1)^(1/9))-1
For 9 years of growth from B1 to B10
