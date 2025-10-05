### 183. BI: How do you calculate year-to-date (YTD) metrics?

Year-to-date (YTD) analysis is a staple of Business Intelligence (BI), measuring performance from the beginning of the current year (either calendar or fiscal) to the present day. These formulas allow you to create dynamic YTD calculations that update automatically.

For these examples, assume you have named ranges `Sales` for your values and `Dates` for the corresponding dates.

#### YTD Sum

This formula calculates the total sum of a value (e.g., sales) from January 1st of the current year up to today.

**Classic Formula:**
```excel
=SUMIFS(Sales, Dates, ">="&DATE(YEAR(TODAY()),1,1), Dates, "<="&TODAY())
```

**How it works:**
*   `DATE(YEAR(TODAY()),1,1)`: This dynamically constructs the start date. `YEAR(TODAY())` gets the current year, and `DATE()` builds the full date for January 1st of that year.
*   `">="&...`: The first criterion filters for dates on or after January 1st of the current year.
*   `"<="&TODAY()`: The second criterion filters for dates on or before today's date.
*   `SUMIFS`: Sums the `Sales` range only for rows that meet both date criteria.

#### YTD Average

This is identical to the YTD Sum, but it calculates the average instead.

**Classic Formula:**
```excel
=AVERAGEIFS(Sales, Dates, ">="&DATE(YEAR(TODAY()),1,1), Dates, "<="&TODAY())
```
The conditions work exactly as in the `SUMIFS` example, but the `AVERAGEIFS` function is used for the aggregation.

#### YTD vs. Prior YTD Comparison

To add context, you can compare the current YTD performance with the performance over the same exact period in the prior year.

**Formula:**
```excel
=(Current_YTD - Prior_YTD) / Prior_YTD
```

> [!IMPORTANT]
> To get the `Prior_YTD` value, you must adjust the date criteria in your `SUMIFS` formula to reference the previous year. A reliable way to do this is using the `EDATE` function to find the equivalent date one year ago.
>
> **Prior YTD Sum Formula:**
> `=SUMIFS(Sales, Dates, ">="&DATE(YEAR(TODAY())-1,1,1), Dates, "<="&EDATE(TODAY(),-12))`

#### YTD with a Fiscal Year

If your organization uses a fiscal year that doesn't start on January 1st, you can adjust the start date accordingly.

**Formula:**
```excel
=SUMIFS(Sales, Dates, ">="&Fiscal_Year_Start, Dates, "<="&TODAY())
```
> [!TIP]
> The best practice is to put the start date of the current fiscal year (e.g., "10/1/2023") into a dedicated cell and name that cell `Fiscal_Year_Start`. This makes your formula easy to update each year without having to edit the formula itself.

#### Dynamic YTD (Excel 365)

The `FILTER` function in Excel 365 provides a more modern and arguably more readable way to perform YTD calculations.

**Formula:**
```excel
=SUM(FILTER(Sales, (YEAR(Dates)=YEAR(TODAY()))*(Dates<=TODAY())))
```
**How it works:**
1.  `FILTER(Sales, ...)`: This function returns only the sales figures that meet the specified conditions.
2.  `(YEAR(Dates)=YEAR(TODAY()))`: This is the first condition. It creates a TRUE/FALSE array for all dates that fall within the current year.
3.  `(Dates<=TODAY())`: This is the second condition, which identifies all dates on or before today.
4.  `*`: The asterisk acts as an `AND` operator, ensuring that only rows where **both** conditions are TRUE are included in the filtered result.
5.  `SUM(...)`: Finally, `SUM` calculates the total of the filtered sales values.
