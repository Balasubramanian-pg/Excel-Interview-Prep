### 125. **How do you use array logic for complex conditions?**

**Count rows meeting all of multiple conditions:**
=SUMPRODUCT((A1:A100="X")*(B1:B100>50)*(C1:C100<100)*(D1:D100="Active"))

**Count with date ranges:**
=SUMPRODUCT((A1:A100>=StartDate)*(A1:A100<=EndDate)*(B1:B100="Product"))

**Sum with complex OR logic:**
=SUMPRODUCT(((A1:A100="West")+(A1:A100="East")>0)*B1:B100)
