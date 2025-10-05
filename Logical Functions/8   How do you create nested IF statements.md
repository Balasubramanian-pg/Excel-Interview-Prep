### 8. **How do you create nested IF statements?**

Multiple IF functions inside each other:

```
=IF(A1>=90, "A", IF(A1>=80, "B", IF(A1>=70, "C", IF(A1>=60, "D", "F"))))

```

Best practices:

- Keep nesting to 3-4 levels maximum for readability
- Consider using IFS() instead for multiple conditions
- Use proper indentation when writing complex formulas
