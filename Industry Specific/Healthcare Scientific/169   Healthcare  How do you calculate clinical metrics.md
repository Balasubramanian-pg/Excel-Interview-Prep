### 169. **Healthcare: How do you calculate clinical metrics?**

**Length of Stay (LOS):**
=Discharge_Date - Admission_Date

**Average LOS:**
=AVERAGE(LOS_Range)

**Readmission Rate:**
=COUNTIFS(Readmission_Flag, "Yes", Days_Since_Discharge, "<=30") / Total_Discharges

**Bed Occupancy Rate:**
=(Patient_Days / (Available_Beds * Days_in_Period)) * 100

**Bed Turnover Rate:**
=Admissions / Average_Number_of_Beds

**Case Mix Index:**
=SUM(DRG_Weights) / Total_Discharges
