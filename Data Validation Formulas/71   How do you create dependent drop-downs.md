### 71. **How do you create dependent drop-downs?**

**Step 1:** Name ranges for each category
**Step 2:** First dropdown uses list of categories
**Step 3:** Second dropdown uses: =INDIRECT($A1)

Where A1 contains the selected category name

**Without named ranges (Excel 365):**
=FILTER(ProductList, CategoryList=A1)
