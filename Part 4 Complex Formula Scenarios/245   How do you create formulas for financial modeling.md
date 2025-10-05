### 245. **How do you create formulas for financial modeling?**

```
Sub CreateFinancialModelFormulas()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Revenue model with growth
    ws.Range("B2").Formula = "=B1*(1+$Growth)"  ' Assuming Growth is named range
    ws.Range("B2").AutoFill Destination:=ws.Range("B2:B11")

    ' COGS as percentage of revenue
    ws.Range("C2:C11").Formula = "=B2*$COGS_Percentage"

    ' Gross Profit
    ws.Range("D2:D11").Formula = "=B2-C2"

    ' Operating Expenses (multiple categories)
    ws.Range("E2:E11").Formula = "=B2*$OpEx_Percentage"

    ' EBITDA
    ws.Range("F2:F11").Formula = "=D2-E2"

    ' Depreciation (straight-line)
    ws.Range("G2").Formula = "=$CapEx/$Useful_Life"
    ws.Range("G2").AutoFill Destination:=ws.Range("G2:G11")

    ' EBIT
    ws.Range("H2:H11").Formula = "=F2-G2"

    ' Interest Expense
    ws.Range("I2:I11").Formula = "=$Debt*$Interest_Rate"

    ' EBT
    ws.Range("J2:J11").Formula = "=H2-I2"

    ' Tax
    ws.Range("K2:K11").Formula = "=J2*$Tax_Rate"

    ' Net Income
    ws.Range("L2:L11").Formula = "=J2-K2"

    ' Free Cash Flow
    ws.Range("M2:M11").Formula = "=F2-$CapEx+G2-$Working_Capital_Change"

    ' NPV Calculation
    ws.Range("N2").Formula = "=NPV($Discount_Rate,M2:M11)+M1"

    ' IRR Calculation
    ws.Range("O2").Formula = "=IRR(M1:M11)"
End Sub

```

**DCF Model with Sensitivity Analysis:**

```
Sub CreateDCFModel()
    Dim ws As Worksheet
    Set ws = Worksheets.Add
    ws.Name = "DCF Model"

    ' Headers
    ws.Range("A1:F1").Value = Array("Year", "Revenue", "EBITDA", "FCF", "PV of FCF", "Terminal Value")

    ' Year numbers
    Dim i As Long
    For i = 1 To 10
        ws.Cells(i + 1, 1).Value = i
    Next i

    ' Revenue projection
    ws.Range("B2").Formula = "=$Starting_Revenue"
    ws.Range("B3:B11").Formula = "=B2*(1+$Revenue_Growth)"

    ' EBITDA
    ws.Range("C2:C11").Formula = "=B2*$EBITDA_Margin"

    ' Free Cash Flow
    ws.Range("D2:D11").Formula = "=C2*(1-$Tax_Rate)-$CapEx-$NWC_Change"

    ' Present Value of FCF
    ws.Range("E2").Formula = "=D2/((1+$WACC)^A2)"
    ws.Range("E2").AutoFill Destination:=ws.Range("E2:E11")

    ' Terminal Value (in final year)
    ws.Range("F11").Formula = "=(D11*(1+$Terminal_Growth))/($WACC-$Terminal_Growth)/((1+$WACC)^10)"

    ' Enterprise Value
    ws.Range("B13").Value = "Enterprise Value"
    ws.Range("C13").Formula = "=SUM(E2:E11)+F11"

    ' Equity Value
    ws.Range("B14").Value = "Equity Value"
    ws.Range("C14").Formula = "=C13-$Debt+$Cash"

    ' Share Price
    ws.Range("B15").Value = "Share Price"
    ws.Range("C15").Formula = "=C14/$Shares_Outstanding"

    ' Create sensitivity table for WACC and Terminal Growth
    CreateSensitivityTable ws
End Sub

Sub CreateSensitivityTable(ws As Worksheet)
    ' Sensitivity analysis table
    ws.Range("H1").Value = "Sensitivity: Share Price"
    ws.Range("I1").Value = "Terminal Growth →"
    ws.Range("H2").Value = "WACC ↓"

    ' Terminal growth rates across top
    Dim tgRates As Variant
    tgRates = Array(0.02, 0.025, 0.03, 0.035, 0.04)
    ws.Range("J1:N1").Value = tgRates

    ' WACC rates down side
    Dim waccRates As Variant
    waccRates = Array(0.08, 0.09, 0.1, 0.11, 0.12)
    ws.Range("I2:I6").Value = Application.Transpose(waccRates)

    ' Formula for sensitivity table
    Dim r As Long, c As Long
    For r = 2 To 6
        For c = 10 To 14
            ws.Cells(r, c).Formula = _
                "=DCF_SharePrice($I" & r & "," & ws.Cells(1, c).Address & ")"
        Next c
    Next r
End Sub

```
