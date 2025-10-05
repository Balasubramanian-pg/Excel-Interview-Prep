### 151. **Finance: How do you calculate loan amortization schedules?**

**Monthly Payment:**
=PMT(Annual_Rate/12, Years*12, -Loan_Amount)

**Detailed amortization table:**

- **Payment Number:** =ROW()-1
- **Beginning Balance:** =IF(ROW()=2, Loan_Amount, Previous_Ending_Balance)
- **Payment:** =$B$1 (absolute reference to PMT formula)
- **Interest:** =Beginning_Balance * (Annual_Rate/12)
- **Principal:** =Payment - Interest
- **Ending Balance:** =Beginning_Balance - Principal

**Total Interest Paid:**
=CUMIPMT(rate/12, nper*12, pv, start_period, end_period, 0)

**Total Principal Paid:**
=CUMPRINC(rate/12, nper*12, pv, start_period, end_period, 0)
