### 48. **How do you create a loan amortization schedule?**

1. PMT for monthly payment: =PMT(rate/12, months, -loan_amount)
2. For each month:
    - Interest: =IPMT(rate/12, month_num, total_months, -loan_amount)
    - Principal: =PPMT(rate/12, month_num, total_months, -loan_amount)
    - Balance: =Previous_Balance - Principal_Payment
