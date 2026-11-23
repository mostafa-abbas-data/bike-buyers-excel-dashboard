# Bike Buyers: Data Cleaning & Dashboard in Excel

An Excel project that cleans a bike buyers dataset, builds pivot tables, and presents the results in an interactive dashboard with slicers.

**Tools:** Microsoft Excel (formulas, pivot tables, pivot charts, slicers)

---

## Project Goals

1. Clean the raw customer data so it is consistent and free of duplicates.
2. Create a customer age group column.
3. Find out which customer characteristics are linked to buying a bike, and show them in a dashboard.

## Files

| File | Description |
|---|---|
| `bike_buyers_before_cleaning.xlsx` | Raw data: 1,026 rows, 13 columns |
| `bike_buyers_after_cleaning.xlsx` | Cleaned data: 1,000 rows, 14 columns |
| `bike_buyers_project_with_dashboard.xlsx` | Full workbook: raw data, cleaned data, pivot tables, and the dashboard |

## Dataset

One row per customer.

| Column | Description |
|---|---|
| ID | Customer ID |
| Marital Status | Married / Single |
| Gender | Male / Female |
| Income | Yearly income |
| Children | Number of children (0 to 5) |
| Education | Education level |
| Occupation | Job category |
| Home Owner | Yes / No |
| Cars | Number of cars owned (0 to 4) |
| Commute Distance | Distance to work, in groups |
| Region | North America / Europe / Pacific |
| Age | Customer age |
| Age-Brackets | Age group (added during cleaning) |
| Purchased Bike | Yes / No |

https://www.youtube.com/redirect?event=video_description&redir_token=QUlxb01IYmdoeHVMV3Y2cW5qNzBnclR6SGRMVXxBQnJ0NjluMXRWZThYdDBPd21GN213Z2hmajY2dGVYRXBuZGJQQkw4c21hY045N1JsRGRiTnlyd0RLTTJyOWR1ZnZJcXRZQTFDdjh2RTI2QWJPTzNUU1MxUEpINGgwbEdfTk5r&q=https%3A%2F%2Fgithub.com%2FAlexTheAnalyst%2FExcel-Tutorial%2Fblob%2Fmain%2FExcel%2520Project%2520Dataset.xlsx&v=opJgMj1IUrc
## Data Cleaning

Comparing the raw and cleaned data shows exactly what changed:

| Problem | Fix | Rows affected |
|---|---|---|
| Duplicate customers (identical rows with the same ID) | Removed | 26 rows (1,026 down to 1,000) |
| Marital Status written as `M` / `S` | Changed to `Married` / `Single` | all rows |
| Gender written as `M` / `F` | Changed to `Male` / `Female` | all rows |
| Commute Distance label `10+ Miles` | Changed to `More than 10 miles` | all rows with this label |
| No age grouping | Added `Age-Brackets` column with an `IF` formula | all rows |

The raw data had no blank cells, so no missing-value handling was needed.

**Age brackets rule**

```excel
=IF(L2>54,"Old",IF(L2>=31,"Middle Age",IF(L2<31,"younger ones ","invalid")))
```

| Age | Bracket |
|---|---|
| 55 and above | Old |
| 31 to 54 | Middle Age |
| Below 31 | younger ones |

After cleaning, the table has 1,000 unique customer IDs and every other column is unchanged.

## Dashboard

The dashboard sheet shows three charts built from pivot tables, with slicers for **Marital Status**, **Education**, and **Region**:

- Average income by gender and purchase decision
- Customers by age bracket and purchase decision
- Customers by commute distance and purchase decision

<!-- Add a screenshot of the dashboard: save it as dashboard.png in this folder, then replace this line with: ![Dashboard](dashboard.png) -->

## Key Findings

Out of 1,000 customers, 481 bought a bike (48.1%).

**Income**
- Buyers earn more on average than non-buyers: about 57,963 vs 54,875.
- Men have a higher average income than women (about 58,063 vs 54,581). Among buyers the gap is similar (about 60,124 vs 55,774).

**Age**
- Middle Age customers (31 to 54) are the biggest group (701 customers) and have the highest purchase rate (54.6%).
- Old customers (55 and above) bought at a much lower rate (31.2%), and so did the youngest group (35.5%).

**Commute distance**

| Commute distance | Customers | Bought | Purchase rate |
|---|---|---|---|
| 0-1 Miles | 366 | 200 | 54.6% |
| 1-2 Miles | 169 | 77 | 45.6% |
| 2-5 Miles | 162 | 95 | 58.6% |
| 5-10 Miles | 192 | 76 | 39.6% |
| More than 10 miles | 111 | 33 | 29.7% |

- The purchase rate is highest for commutes of 2-5 miles (58.6%) and 0-1 miles (54.6%), and lowest for long commutes: 39.6% for 5-10 miles and 29.7% for more than 10 miles.

These are patterns in this sample, not proof of cause and effect. Some groups are small (for example, 110 customers in the youngest bracket).

## Notes and Possible Improvements

- The youngest age group label is `younger ones ` (lowercase, with a trailing space). Renaming it to something like `Young` would make the labels consistent; the pivot tables update after a refresh.
- Eleven customers are older than 70 (the oldest is 89). They were kept because the values are plausible.
- To explore the dashboard, open `bike_buyers_project_with_dashboard.xlsx` in Excel and use the slicers.
