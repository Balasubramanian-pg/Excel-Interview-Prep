### 75. **How do you check if text contains a specific pattern?**

Here are several ways to check if text contains specific patterns in Excel:

## **1. Contains Specific Text**
```excel
=ISNUMBER(SEARCH("text", A1))
=ISNUMBER(FIND("text", A1))
```
*Note: SEARCH is case-insensitive, FIND is case-sensitive*

## **2. Starts with Specific Text**
```excel
=LEFT(A1, LEN("text"))="text"
=COUNTIF(A1, "text*")>0
```

## **3. Ends with Specific Text**
```excel
=RIGHT(A1, LEN("text"))="text"
=COUNTIF(A1, "*text")>0
```

## **4. Using Wildcards with COUNTIF**
```excel
=COUNTIF(A1, "*pattern*")>0          ' Contains pattern
=COUNTIF(A1, "start*")>0             ' Starts with
=COUNTIF(A1, "*end")>0               ' Ends with
=COUNTIF(A1, "a??b")>0               ' Specific pattern with wildcards
```

## **5. Regular Expressions (with VBA)**
For more complex patterns, create a custom function:
```vba
Function RegExMatch(cell As Range, pattern As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.pattern = pattern
    regEx.IgnoreCase = True
    RegExMatch = regEx.Test(cell.Value)
End Function
```
Usage: `=RegExMatch(A1, "\d{3}-\d{2}-\d{4}")` (matches SSN pattern)

## **6. Common Pattern Examples**
```excel
=COUNTIF(A1, "???-??-????")>0        ' SSN-like pattern
=COUNTIF(A1, "[A-Z]*[0-9]")>0        ' Starts with letter, ends with number
=AND(ISNUMBER(SEARCH("@",A1)), ISNUMBER(SEARCH(".",A1)))  ' Basic email check
```

## **7. Multiple Conditions**
```excel
=AND(COUNTIF(A1, "*text1*")>0, COUNTIF(A1, "*text2*")>0)
=OR(COUNTIF(A1, "*option1*")>0, COUNTIF(A1, "*option2*")>0)
```

The **COUNTIF with wildcards** approach is often the most readable for simple pattern matching, while **VBA RegEx** provides the most powerful pattern matching capabilities for complex requirements.

