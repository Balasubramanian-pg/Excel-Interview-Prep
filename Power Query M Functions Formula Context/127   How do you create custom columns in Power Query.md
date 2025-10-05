### 127. **How do you create custom columns in Power Query?**

```
= if [Sales] > 1000 then "High" else "Low"

```

**Nested conditions:**

```
= if [Value] > 100 then "A"
  else if [Value] > 50 then "B"
  else "C"

```

**Multiple column logic:**

```
= if [Status] = "Active" and [Amount] > 1000 then "Priority" else "Standard"

```
