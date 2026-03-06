# Multi Filter Stock Analysis

## Overview

This project performs **multi-filter stock analysis** using valuation metrics, growth filters, and debt filters to identify high-quality investment opportunities.
The script processes financial datasets, applies multiple conditions, and generates structured Excel reports for analysis.

The analysis focuses on combining **valuation metrics (PS, PB, MCAP/OCF)** with **growth indicators (Revenue Growth, PAT Growth)** to identify strong stocks across different market cap and debt scenarios.

---

## Features

* Multi-filter stock screening
* Valuation analysis using:

  * Price-to-Sales (PS)
  * Price-to-Book (PB)
  * MCAP/OCF ratio
* Growth-based filtering:

  * Revenue Growth
  * PAT Growth
* Net Debt filtering for safer portfolio picks
* Forward return and CAGR analysis
* Automated Excel report generation
* Portfolio stock selection based on combined filters

---

## Technologies Used

* Python
* Pandas
* NumPy
* XlsxWriter
* Excel Data Processing

---

## Project Workflow

1. Load financial dataset from Excel
2. Clean and preprocess the data
3. Apply sector filtering (exclude Bank and Insurance)
4. Apply valuation filters (PS / PB bins)
5. Apply growth filters (Revenue & PAT growth)
6. Apply debt filters (Net Debt conditions)
7. Compute forward returns and CAGR metrics
8. Generate Excel reports for each filter scenario
9. Produce a final **portfolio picks sheet**

---

## Output

The script automatically generates:

* Multiple Excel reports for each filter scenario
* Growth vs Non-Growth analysis tables
* Summary performance metrics
* Selected portfolio picks

Example output structure:

Output_<timestamp>/
│
├── PS_OCF_0_20_NetDebt_LT0_Mar.xlsx
├── PB_OCF_0_50_Debt_NoFilter_Mar.xlsx
└── Master_Summary_<timestamp>.xlsx

---

## Example Portfolio Filters

The script selects stocks based on conditions such as:

* PB ≤ 3 and MCAP/OCF between 0–50
* PB ≤ 5 and MCAP/OCF between 0–20
* Revenue Growth ≥ 15%
* PAT Growth ≥ 15%
* Net Debt ≤ 0.3

These filters help identify **high-growth and financially stable companies**.

---

## How to Run the Project

1. Install dependencies

pip install pandas numpy xlsxwriter

2. Update the dataset path inside the script.

3. Run the script:

python Multi_Filter_Analysis_4_Mar_2026.py

The Excel reports will be generated automatically.

---

## Future Improvements

* Convert analysis into a Streamlit dashboard
* Add automated data ingestion pipeline
* Integrate visualization charts
* Add backtesting capability
* Deploy as an investment research tool

---

## Author

Prathmesh Bondre
Data Science | Machine Learning | Financial Data Analysis
