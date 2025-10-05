### 142. **How do you identify and handle circular references?**

**Intentional iterative calculations:**
Enable: File → Options → Formulas → Enable Iterative Calculation

**Example - iterative convergence:**
=IF(A1="", 100, A1*0.9+10)

Converges to a stable value through iteration

**Prevent circular reference errors:**
Use helper columns or break the circular logic into steps
