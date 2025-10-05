### 197. **How do you create dynamic date ranges for reports?**

**Current Month:**

- Start: =EOMONTH(TODAY(),-1)+1
- End: =EOMONTH(TODAY(),0)

**Last Month:**

- Start: =EOMONTH(TODAY(),-2)+1
- End: =EOMONTH(TODAY(),-1)

**Quarter-to-Date:**

- Start: =DATE(YEAR(TODAY()), CEILING(MONTH(TODAY())/3,1)*3-2, 1)
- End: =TODAY()

**Last N Days:**

- Start: =TODAY()-N
- End: =TODAY()

**Trailing 12 Months:**

- Start: =EDATE(TODAY(),-12)
- End: =TODAY()

**Week Starting Monday:**

- Start: =TODAY()-WEEKDAY(TODAY(),2)+1
- End: =TODAY()-WEEKDAY(TODAY(),2)+7

**Fiscal Year (July start):**
=IF(MONTH(TODAY())>=7, YEAR(TODAY())+1, YEAR(TODAY()))
