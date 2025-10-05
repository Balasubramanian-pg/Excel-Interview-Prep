### 170. **Healthcare: How do you calculate patient satisfaction scores?**

**Net Promoter Score (NPS):**
=Percentage_Promoters - Percentage_Detractors

Where Promoters = 9-10 rating, Detractors = 0-6 rating

**HCAHPS Top Box Score:**
=COUNTIFS(Response, "9", Response, "10") / COUNT(Responses)

**Patient Satisfaction Index:**
=AVERAGE(Satisfaction_Scores) * 100 / Maximum_Score
