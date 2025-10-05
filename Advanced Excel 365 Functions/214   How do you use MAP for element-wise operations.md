### 214. **How do you use MAP for element-wise operations?**

**Apply function to each element:**
=MAP(A1:A10, LAMBDA(x, x^2))

**Multi-array operation:**
=MAP(A1:A10, B1:B10, LAMBDA(a, b, a*b + b^2))

**Conditional transformation:**
=MAP(A1:A10, LAMBDA(x, IF(x>100, x*1.1, x)))

**Text transformation:**
=MAP(A1:A10, LAMBDA(x, PROPER(TRIM(x))))

**Date operations:**
=MAP(Dates, LAMBDA(d, TEXT(d, "mmmm yyyy")))
