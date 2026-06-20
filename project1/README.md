# Project 1: Data Cleaning & Preparation 🧹

**Industrial Training Kit – Batch 2026 | Powered by DecodeLabs**

## 📋 Overview
This project takes a raw e-commerce orders dataset (1,200 records) and transforms it into a clean, production-ready "gold standard" dataset following professional data integrity standards.

## 🎯 Objectives
- Identify and handle missing/null values
- Remove duplicate records
- Correct data formats (dates, numbers, text)
- Document every change made for full accountability

## 📂 Files
- `Dataset_for_Data_Analytics.xlsx` — Original raw dataset (1,200 rows, 14 columns)
- `Dataset_for_Data_Analytics_CLEANED.xlsx` — Cleaned, formatted, and audited dataset

## 🧼 Cleaning Steps Performed

| Change ID | Description | Impact | Status |
|-----------|-------------|--------|--------|
| CR001 | Filled missing `CouponCode` values (309 rows) with `NoCoupon` category | Preserved all 309 records — clarified that blank = no coupon used, not a data gap | ✅ Resolved |
| CR002 | Standardized `Date` column to ISO 8601 (`YYYY-MM-DD`) | All 1,200 dates verified compliant | ✅ Resolved |
| CR003 | Rounded `UnitPrice` and `TotalPrice` to 2 decimal places | Removed floating-point precision artifacts | ✅ Resolved |
| CR004 | Trimmed whitespace & verified consistent text casing across categorical fields | 0 inconsistencies found | ✅ Resolved |
| CR005 | Audited `OrderID` for duplicates | 0 duplicates found across 1,200 rows | ✅ Verified |
| CR006 | Audited full-row duplicates | 0 fully duplicated rows found | ✅ Verified |
| CR007 | Checked `Quantity × UnitPrice = TotalPrice` consistency | 0 mismatches found | ✅ Verified |
| CR008 | Checked for negative/zero `Quantity`, `UnitPrice`, `ItemsInCart` | 0 invalid values found | ✅ Verified |

## ✅ Verification Gate (Project 2 Threshold)

| Metric | Result |
|--------|--------|
| Duplicate OrderID Count | **0** |
| Incorrectly Formatted Dates | **0** |
| Missing Values (after cleaning) | **0** |
| Quantity × UnitPrice = TotalPrice Mismatches | **0** |
| Error Rate on Unique Identifiers | **0%** |
| Error Rate on Date Formats | **0%** |

> *"Before you finish, you must prove there are zero duplicate IDs and zero incorrectly formatted dates."* — Threshold met. ✔️

## 🛠️ Tools Used
- Python (pandas, openpyxl)

## 🚀 Key Skills Demonstrated
Data cleaning, data preparation, strategic imputation, data integrity auditing, documentation & change-log accountability.

---
*Part of the DecodeLabs Industrial Training Kit – Batch 2026*
