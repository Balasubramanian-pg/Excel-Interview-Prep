### 26. **How do you add working days to a date?**

- **WORKDAY(start_date, days, [holidays])**: Adds working days
Example: =WORKDAY(TODAY(), 10, H1:H10) returns date 10 working days from today
- **WORKDAY.INTL(start_date, days, [weekend], [holidays])**: Custom weekends
