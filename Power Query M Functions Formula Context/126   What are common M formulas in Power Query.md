### 126. **What are common M formulas in Power Query?**

**Text operations:**

- Text.Combine: =Text.Combine({"Hello", "World"}, " ")
- Text.BeforeDelimiter: Extract before character
- Text.AfterDelimiter: Extract after character
- Text.BetweenDelimiters: Extract between characters

**Date operations:**

- Date.AddDays: =Date.AddDays(#date(2025,1,1), 30)
- Date.DayOfWeek: Returns day (0-6)
- Date.StartOfMonth: First day of month

**List operations:**

- List.Sum, List.Average, List.Max
- List.Distinct: Unique values
- List.Select: Filter list
