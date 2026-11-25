# Global Electronics Retailer — Sales Performance Analysis

## Project overview

This Excel portfolio project analyzes a multi-year global electronics retail dataset. It explores how product list-price sales, estimated gross profit, orders, and units vary by year, country, product category, and month. The interactive dashboard is designed to let a reviewer explore the results with year slicers.

## Dataset

The workbook uses the Global Electronics Retailer CSV tables for sales, products, customers, stores, and exchange rates. The sales table contains 62,884 line items. Each row represents a product line within an order; `Order Number` plus `Line Item` identifies the line, while `Quantity` records units on that line.

## Approach and tools

- Used Power Query to inspect and clean source data types, dates, and relationships between tables. Missing delivery dates were retained rather than filled with invented values.
- Built an Excel Data Model with related sales, product, customer, store, and exchange-rate data.
- Created Power Pivot measures for estimated list-price sales, estimated product gross profit, estimated gross margin, distinct orders, and average list-price value per order.
- Used PivotTables, PivotCharts, and slicers to explore and present results. The dashboard KPI values are displayed from PivotTable-based summaries.
- Practiced Excel formulas separately with `XLOOKUP`, `SUMIFS`, `COUNTIFS`, and `IF` on the Formula Practice sheet.

## Key results

All-years results in the current workbook view:

- **Estimated gross sales at product list price:** $55.76 million
- **Estimated product gross profit at list price:** $32.66 million
- **Estimated gross margin at list price:** 58.6%
- **Distinct orders:** 26,326
- **Average list-price value per order:** $2,117.89
- **Units sold:** 197,757

Additional observations:

- The United States had the highest total list-price sales, about **$29.87 million**, across **14,221 orders**.
- Australia had the highest average list-price value per order among the countries compared, at **$2,291.15**.
- Computers generated the highest list-price sales by category; Games and Toys generated the lowest.
- **2019** was the strongest complete year at about **$18.26 million**. **2021 contains January and February only**, so it is not comparable with full years.
- In the complete years from 2016 through 2020, the monthly view showed a decline from March to April and an increase from November to December. This is a pattern in the data, not an explanation of its cause.

## Important metric definitions and limitations

The sales and profit figures in this project are **estimates based on product list prices**, not recorded transaction revenue or accounting profit. Estimated list-price sales are calculated from quantity multiplied by product unit list price. Estimated product gross profit uses quantity multiplied by the difference between unit list price and unit cost. These measures do not account for discounts, returns, taxes, shipping, or other adjustments, so they should not be presented as actual revenue or realized profit.

The 2021 results cover only January and February. Missing delivery dates were left blank because the source data does not provide a delivery date for most records; no dates were imputed.

## Files

- `Global_Electronics_Sales_Analysis.xlsx` — interactive workbook with the dashboard, formula practice, Pivot analysis, and cleaned worksheet tables.
- `Portfolio Screenshots/Dashboard_Overview.png` — full-dashboard overview screenshot.

## How to explore the workbook

Open the workbook on the **Dashboard** sheet. Use the main **Years** slicer to filter the KPI cards and connected charts. The **YearForMonthlyTrend** slicer controls the monthly trend chart independently, so it can show one selected year without changing the other dashboard visuals.


