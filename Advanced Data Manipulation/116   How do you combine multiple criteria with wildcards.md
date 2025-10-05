### 116. How do you combine multiple criteria with wildcards?

Combining wildcards with multiple criteria allows you to perform flexible and powerful searches within your data, such as counting or summing based on partial text matches across several columns.

### Method 1: Standard Wildcards in `COUNTIFS` and `SUMIFS`

The `...IFS` family of functions fully supports wildcards for criteria, allowing you to filter based on patterns. This method works for conditions that all need to be true at the same time (AND logic).

> [!NOTE]
> **Excel Wildcards:**
> *   `*` (asterisk): Represents any number of characters. `App*` matches "Apple", "Application", etc.
> *   `?` (question mark): Represents any single character. `H?t` matches "Hot", "Hat", "Hit".
> *   `~` (tilde): Acts as an escape character to find a literal `*`, `?`, or `~`. `*~?*` finds text containing a literal question mark.

**Formula:**
This formula counts rows where the text in column A starts with "Apple" AND the text in column B contains the word "Red" anywhere within it.
```excel
=COUNTIFS(A:A, "Apple*", B:B, "*Red*")
```

**How it works:**
*   `A:A, "Apple*"`: The first criterion checks for any text in column A that begins with "Apple".
*   `B:B, "*Red*"`: The second criterion checks for any text in column B that has "Red" anywhere in the cell's content.
*   `COUNTIFS` only counts the row if **both** of these conditions are met for that row.

### Method 2: Complex `OR` Logic with Wildcards using `SUMPRODUCT`

The `...IFS` functions cannot handle an `OR` condition when you are looking for multiple different text patterns within the *same* column. For this, you need to use array logic, and `SUMPRODUCT` is the classic tool for the job.

**Formula:**
This formula sums the values in column B where the corresponding cell in column A contains **either** "keyword1" **or** "keyword2".
```excel
=SUMPRODUCT((ISNUMBER(SEARCH("keyword1", A:A)) + ISNUMBER(SEARCH("keyword2", A:A)) > 0) * B:B)
```

**How it works:**
This formula is a multi-step array calculation:
1.  `SEARCH("keyword1", A:A)`: This searches for "keyword1" in column A. It returns a number (the start position) if found, and a `#VALUE!` error if not.
2.  `ISNUMBER(...)`: This converts the result from `SEARCH` into a `TRUE`/`FALSE` array. You get `TRUE` if the keyword was found (a number) and `FALSE` if it wasn't (an error).
3.  `(...) + (...)`: The `+` operator acts as a logical `OR`. We perform steps 1 and 2 for both keywords. By adding the two `TRUE`/`FALSE` arrays, any row where at least one keyword was found will have a value of `1` or greater (`TRUE`=`1`, `FALSE`=`0`).
4.  `> 0`: This step converts the array of numbers (`{1; 0; 2; ...}`) back into a final `TRUE`/`FALSE` array that identifies all rows matching at least one of the keywords.
5.  `* B:B`: The final `TRUE`/`FALSE` array is multiplied by the values in column B. The values in rows that did not match are multiplied by `0` (becoming zero), and the values in rows that did match are multiplied by `1` (remaining unchanged).
6.  `SUMPRODUCT` then sums the final array of values.

> [!IMPORTANT]
> In `SUMPRODUCT` array logic:
> *   Multiplication `(*)` acts as an **AND** operator.
> *   Addition `(+)` acts as an **OR** operator.

> [!TIP]
> **Modern Excel 365 Alternative with `FILTER`**
> The `FILTER` function can accomplish the `OR` logic in a much more readable way, though it requires an extra step to sum the result.
> ```excel
> =SUM(FILTER(B:B, (ISNUMBER(SEARCH("keyword1", A:A))) + (ISNUMBER(SEARCH("keyword2", A:A)))))
> ```
