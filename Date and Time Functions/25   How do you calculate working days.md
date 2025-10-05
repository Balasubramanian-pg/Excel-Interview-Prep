### 25. **How do you calculate working days?**

- **NETWORKDAYS(start_date, end_date, [holidays])**: Working days excluding weekends
Example: =NETWORKDAYS(A1, B1, H1:H10) excludes weekends and holidays in H1:H10
- **NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])**: Custom weekends
Example: =NETWORKDAYS.INTL(A1, B1, 7) treats only Sunday as weekend
