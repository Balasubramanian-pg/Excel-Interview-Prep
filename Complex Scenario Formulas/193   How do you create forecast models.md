### 193. **How do you create forecast models?**

**Linear Trend Forecast:**
=FORECAST.LINEAR(New_X, Known_Y's, Known_X's)

**Seasonal forecast:**
=FORECAST.ETS(target_date, values, timeline, [seasonality], [data_completion], [aggregation])

**Growth trend:**
=GROWTH(known_y's, known_x's, new_x's, [const])

**Exponential smoothing:**
=FORECAST.ETS.STAT(values, timeline, statistic_type, [seasonality], [data_completion], [aggregation])
