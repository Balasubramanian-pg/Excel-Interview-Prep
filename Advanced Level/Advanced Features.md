### Advanced Features

**Macros:**
Recorded or written VBA code that automates tasks.

1. View tab → Macros → Record Macro
2. Perform your actions
3. Stop recording
4. Edit via View → Macros → View Macros → Edit (opens VBA editor)

**Power Query and Power Pivot:**

- **Power Query**: ETL tool (Extract, Transform, Load) for cleaning and shaping data from multiple sources. Access via Data tab → Get Data
- **Power Pivot**: Advanced data modeling tool for creating relationships between tables, DAX formulas, and handling millions of rows

**Dynamic Named Ranges:**
Named ranges that expand/contract automatically.
Example: =OFFSET($A$1,0,0,COUNTA($A:$A),1)
Creates a range from A1 that includes all non-empty cells below

**What-If Analysis:**

- **Goal Seek**: Find input value needed for desired output (Data tab → What-If Analysis)
- **Scenario Manager**: Save different input sets and compare outcomes
- **Data Tables**: Show how changing 1-2 variables affects formulas

**Protect Worksheets/Workbooks:**

- Review tab → Protect Sheet (locks cells, can set password, allow specific actions)
- Review tab → Protect Workbook (prevents structure changes)
- Unprotect specific cells: Format Cells → Protection tab → Uncheck "Locked" before protecting
