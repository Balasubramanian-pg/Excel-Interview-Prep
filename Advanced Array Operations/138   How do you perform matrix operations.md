### 138. How do you perform matrix operations?

Excel includes a powerful set of functions for performing linear algebra, specifically for matrix operations. These are array functions, meaning they can operate on and return arrays of values, making them ideal for solving complex mathematical problems.

#### Matrix Multiplication (`MMULT`)

The `MMULT` function returns the matrix product of two arrays. The resulting matrix will have the same number of rows as the first array and the same number of columns as the second array.

**Formula:**
```excel
=MMULT(array1, array2)
```
*   `array1`: The first matrix in the multiplication.
*   `array2`: The second matrix in the multiplication.

**Example:**
```excel
=MMULT(A1:C3, E1:G3)
```
This formula multiplies a 3x3 matrix in `A1:C3` by a 3x3 matrix in `E1:G3`.

> [!IMPORTANT]
> To multiply matrices, the number of **columns** in the first matrix (`array1`) must be equal to the number of **rows** in the second matrix (`array2`). If this rule is not met, the formula will return a `#VALUE!` error.

#### Matrix Determinant (`MDETERM`)

The `MDETERM` function calculates the determinant of a matrix, which is a scalar value useful in analyzing and solving systems of linear equations.

**Formula:**
```excel
=MDETERM(array)
```
*   `array`: The matrix for which you want to calculate the determinant.

> [!CAUTION]
> The `MDETERM` function requires a **square matrix** (i.e., the number of rows and columns must be equal). Providing a non-square matrix will result in a `#VALUE!` error.

#### Matrix Inverse (`MINVERSE`)

The `MINVERSE` function calculates the inverse of a matrix. The inverse is essential for solving systems of linear equations.

**Formula:**
```excel
=MINVERSE(array)
```
*   `array`: The square matrix you want to invert.

> [!WARNING]
> A matrix can only be inverted if its determinant is non-zero. If you attempt to use `MINVERSE` on a singular matrix (one with a determinant of 0), the function will return the `#NUM!` error. Like `MDETERM`, this function also requires a square matrix.

#### Practical Application: Solving Systems of Linear Equations

You can combine these functions to solve a system of linear equations (e.g., `Ax = b`) for the variable vector `x` using the formula `x = A⁻¹b`.

**Formula:**
```excel
=MMULT(MINVERSE(coefficients), constants)
```

**Example:**
Consider the following system of equations:
`2x + 3y = 8`
`4x + 1y = 6`

1.  **Set up your data in Excel:**
    *   In `A1:B2`, enter the coefficient matrix:
        *   A1: `2`, B1: `3`
        *   A2: `4`, B2: `1`
    *   In `D1:D2`, enter the constants vector:
        *   D1: `8`
        *   D2: `6`

2.  **Enter the formula in an empty cell:**
    ```excel
    =MMULT(MINVERSE(A1:B2), D1:D2)
    ```

3.  **Result:** The formula will spill a 2x1 array containing the values for `x` and `y`.
    *   `1` (the value for x)
    *   `2` (the value for y)

> [!TIP]
> Before dynamic arrays were introduced in Excel 365, these formulas had to be entered as legacy array formulas by selecting the output range first and pressing `Ctrl+Shift+Enter`. In modern Excel, you just enter the formula in one cell, and it spills the results automatically.
