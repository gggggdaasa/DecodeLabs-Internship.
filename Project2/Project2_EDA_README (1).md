# Project 2 — Exploratory Data Analysis (EDA)

**DecodeLabs Industrial Training Program — Batch 2026**
**Intern:** Mustafa Mahmoud Ali

## Goal

Analyze the e-commerce orders dataset to understand patterns, trends, and distributions in the data.

## Key Requirements

- Calculate basic statistics (mean, median, count, and more)
- Identify trends and outliers
- Summarize key observations

## Dataset

`data/Dataset_for_Data_Analytics.xlsx` — 1,200 e-commerce orders (Jan 2023 – Jun 2025), 14 fields:
`OrderID, Date, CustomerID, Product, Quantity, UnitPrice, ShippingAddress, PaymentMethod, OrderStatus, TrackingNumber, ItemsInCart, CouponCode, ReferralSource, TotalPrice`

## Approach

1. **Load & validate** — checked for duplicate rows/IDs, missing values, and confirmed `TotalPrice = Quantity × UnitPrice` holds across all 1,200 rows.
2. **Descriptive statistics** — mean, median, mode, standard deviation, min/max, and quartiles for all four numeric fields (`Quantity`, `UnitPrice`, `ItemsInCart`, `TotalPrice`).
3. **Group summaries (pivot tables)** — orders/revenue broken down by Product, Order Status, Payment Method, Referral Source, Coupon Code, and Month.
4. **Outlier detection** — IQR method (1.5× rule) applied to `TotalPrice` and `UnitPrice`.
5. **Visualization** — bar charts for category comparisons, a line chart for the monthly revenue trend, and a histogram for the `TotalPrice` distribution.
6. **Key observations** — written summary of findings and a short list of recommended next steps.

## How to run

```bash
pip install pandas numpy openpyxl
python clean_eda.py
```

This regenerates `output/DecodeLabs_Project2_EDA.xlsx` from the raw dataset in `data/`. All statistics in the workbook are live Excel formulas (`AVERAGE`, `MEDIAN`, `MODE.SNGL`, `STDEV.S`, `QUARTILE.INC`, `SUMIF`, `COUNTIF`, …) referencing the Raw Data sheet — nothing is hardcoded, so the workbook recalculates automatically if the source data changes.

**Recalculating formulas:** openpyxl writes formulas but not their cached results. After running the script, open the file in Excel (it will calculate automatically), or recalculate headlessly with LibreOffice:

```bash
soffice --headless --convert-to xlsx --outdir output output/DecodeLabs_Project2_EDA.xlsx
```

## Output

`output/DecodeLabs_Project2_EDA.xlsx` — a single workbook with 6 sheets:

| Sheet | Contents |
|---|---|
| **Key Observations** | Narrative summary: dataset overview, central tendency, trends, order status patterns, outliers, recommendations |
| **Summary Statistics** | Mean/median/mode/std-dev/min/max/quartiles for all numeric fields + IQR outlier bounds |
| **Pivot Tables** | Order count, total revenue, and avg order value grouped by Product, Status, Payment Method, Referral Source, Coupon, and Month |
| **Charts** | 6 charts: revenue by product, orders by status, orders by payment method, avg order value by referral source, monthly revenue trend, TotalPrice histogram |
| **Outliers** | Full list of the 8 orders flagged by the IQR method, with explanation |
| **Raw Data** | Original dataset as a structured Excel Table |

## Key Findings

- **Clean data**: 0 duplicate rows, 0 duplicate OrderIDs. The only missing values are in `CouponCode` (309/1,200 orders) — meaning no coupon was used, not a data quality issue.
- **Right-skewed spending**: average order value is **$1,053.97** vs. a median of **$823.62** — a handful of high-value orders pull the mean up.
- **Balanced catalog**: revenue is fairly evenly spread across all 7 products, 5 payment methods, and 5 referral sources — no single category dominates.
- **41.4% of orders are Cancelled or Returned** — worth investigating further as a possible operational issue.
- **8 outliers (0.7%)** in `TotalPrice`, all genuine large bulk orders (Quantity = 5 with high unit price), not data errors. Zero outliers in `UnitPrice`.
- Coupon usage is common (74.2% of orders) but doesn't meaningfully increase average order value.

## Tools & Skills Used

`pandas`, `numpy`, `openpyxl` — data cleaning, descriptive statistics, group-by aggregation, IQR outlier detection, Excel formula generation, and native Excel chart creation.
