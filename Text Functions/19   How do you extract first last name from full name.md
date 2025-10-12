# 19. **How do you extract first/last name from full name?**

**Extracting First, Middle, and Last Name from Full Name in Excel**

**Extracting First Name**
**Syntax:** `=LEFT(full_name, FIND(" ", full_name)-1)`
**Purpose:** Extracts the first name from a full name string.
**Example:** `=LEFT(A1, FIND(" ", A1)-1)`
**Input:** `A1 = "John William Doe"`
**Output:** `"John"`

> [!NOTE]
> Assumes the full name has at least a first name and a last name separated by a space.

**Extracting Last Name**
**Syntax:** `=RIGHT(full_name, LEN(full_name)-FIND("~", SUBSTITUTE(full_name, " ", "~", LEN(full_name)-LEN(SUBSTITUTE(full_name, " ", "")))))`
**Simpler Alternative:** `=TRIM(RIGHT(SUBSTITUTE(full_name, " ", REPT(" ", 100)), 100))`
**Purpose:** Extracts the last name from a full name string.
**Example:** `=TRIM(RIGHT(SUBSTITUTE(A1, " ", REPT(" ", 100)), 100))`
**Input:** `A1 = "John William Doe"`
**Output:** `"Doe"`

> [!TIP]
> The simpler alternative works by replacing spaces with a large number of spaces and then extracting the last 100 characters, which will be the last name.

**Extracting Middle Name**
**Syntax:** `=TRIM(MID(full_name, FIND(" ", full_name), FIND(" ", full_name, FIND(" ", full_name)+1)-FIND(" ", full_name)))`
**Purpose:** Extracts the middle name from a full name string.
**Example:** `=TRIM(MID(A1, FIND(" ", A1), FIND(" ", A1, FIND(" ", A1)+1)-FIND(" ", A1)))`
**Input:** `A1 = "John William Doe"`
**Output:** `"William"`

> [!WARNING]
> This formula assumes there is exactly one middle name. If there are more than two spaces, it will only extract the text between the first and second space.

**Practical Example**

**Input:** `A1 = "Mary Anne Smith"`
**Goal:** Extract first, middle, and last names.

**First Name:**
**Formula:** `=LEFT(A1, FIND(" ", A1)-1)`
**Output:** `"Mary"`

**Middle Name:**
**Formula:** `=TRIM(MID(A1, FIND(" ", A1), FIND(" ", A1, FIND(" ", A1)+1)-FIND(" ", A1)))`
**Output:** `"Anne"`

**Last Name:**
**Formula:** `=TRIM(RIGHT(SUBSTITUTE(A1, " ", REPT(" ", 100)), 100))`
**Output:** `"Smith"`

**Flashcards: Q&A**

**Q1: How do you extract the first name from a full name in Excel?**
**A:** `=LEFT(A1, FIND(" ", A1)-1)`

**Q2: How do you extract the last name from a full name in Excel?**
**A:** `=TRIM(RIGHT(SUBSTITUTE(A1, " ", REPT(" ", 100)), 100))`

**Q3: How do you extract the middle name from a full name in Excel?**
**A:** `=TRIM(MID(A1, FIND(" ", A1), FIND(" ", A1, FIND(" ", A1)+1)-FIND(" ", A1)))`

**Q4: What does the `TRIM` function do in the context of extracting names?**
**A:** Removes extra spaces to ensure clean output.

**Q5: What assumption does the middle name extraction formula make?**
**A:** It assumes there is exactly one middle name and at least one space before and after it.

> [!IMPORTANT]
> Always validate your data to ensure the formulas work as expected. Use `IFERROR` to handle potential errors if the format of the full name is inconsistent.
