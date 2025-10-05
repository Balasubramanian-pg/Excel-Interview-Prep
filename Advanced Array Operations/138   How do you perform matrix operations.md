### 138. **How do you perform matrix operations?**

**Matrix multiplication:**
=MMULT(array1, array2)

Example: =MMULT(A1:C3, E1:G3)
First matrix columns must equal second matrix rows

**Matrix determinant:**
=MDETERM(array)

**Matrix inverse:**
=MINVERSE(array)

**Solving systems of equations:**
=MMULT(MINVERSE(coefficients), constants)
