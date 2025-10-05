### 47. **Explain common financial functions**

- **PMT(rate, nper, pv, [fv], [type])**: Payment for loan
Example: =PMT(5%/12, 30*12, -200000) returns monthly payment on $200k mortgage at 5% for 30 years
- **FV(rate, nper, pmt, [pv], [type])**: Future value
Example: =FV(8%/12, 20*12, -500, 0, 0) future value saving $500/month at 8% for 20 years
- **PV(rate, nper, pmt, [fv], [type])**: Present value
- **RATE(nper, pmt, pv, [fv], [type])**: Interest rate
- **NPER(rate, pmt, pv, [fv], [type])**: Number of periods
- **IPMT(rate, per, nper, pv, [fv], [type])**: Interest portion of payment
- **PPMT(rate, per, nper, pv, [fv], [type])**: Principal portion of payment
- **NPV(rate, value1, value2, ...)**: Net present value
- **IRR(values, [guess])**: Internal rate of return
- **XIRR(values, dates, [guess])**: IRR with irregular periods
