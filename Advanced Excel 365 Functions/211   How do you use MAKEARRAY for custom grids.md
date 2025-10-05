### 211. **How do you use MAKEARRAY for custom grids?**

**Multiplication table:**
=MAKEARRAY(10, 10, LAMBDA(r, c, r*c))

**Custom pattern generator:**
=MAKEARRAY(5, 5, LAMBDA(r, c, IF(r=c, 1, 0)))
Creates identity matrix

**Distance matrix:**
=MAKEARRAY(ROWS(Locations), ROWS(Locations),
LAMBDA(r, c,
SQRT((INDEX(Lat, r)-INDEX(Lat, c))^2 + (INDEX(Lon, r)-INDEX(Lon, c))^2)
)
)

**Conditional grid:**
=MAKEARRAY(Rows, Cols, LAMBDA(r, c, IF(MOD(r+c, 2)=0, "X", "O")))
