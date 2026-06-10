![Python Version](https://img.shields.io/badge/Python-3.10%2B-blue)
![Pandas](https://img.shields.io/badge/pandas-Data_Manipulation-150458)
![License](https://img.shields.io/badge/License-MIT-green)

# NeoStats Retail Data Engineering Pipeline

## Project Overview

This project implements a robust, end-to-end data engineering pipeline designed to process raw retail transactions for ABC Retail Solutions. Built entirely in Python, the pipeline ingests messy, multi-sheet Excel data, performs rigorous data quality transformations, masks sensitive customer information, and structures the final output into a clean Star Schema optimized for immediate ingestion into Business Intelligence tools.

The pipeline was developed as part of the NeoStats Data Engineering Assessment to demonstrate production-grade practices. It ensures data integrity by handling issues like mixed date formats, inconsistent category spellings, missing prices, and payment conflicts. The ultimate deliverable provides business stakeholders with accurate, ready-to-use business KPIs and a flat analytical dataset to power a comprehensive Power BI dashboard.

## Tech Stack
- Python 3.10+, pandas, openpyxl, hashlib
- Power BI Desktop
- Git / GitHub

## Folder Structure

```text
neostats-retail-pipeline/
├── Code/
│   └── retail_pipeline.py            # Main Python data engineering pipeline script
├── Documentation/
│   ├── project_documentation.md      # Detailed technical architecture & project writeup
│   ├── Power_BI_Guide.md             # Guide on importing data to Power BI
│   └── dashboard_instructions.md     # Step-by-step Power BI dashboard build instructions
├── Power BI/
│   └── retail_dashboard.pbix         # The resulting Power BI dashboard (binary)
├── data/
│   ├── USECASE_-_Data_Engineering.xlsx # Raw source data drop (input)
│   ├── retail_final_cleaned.csv        # Flattened analytical dataset (output)
│   ├── powerbi_dataset.xlsx            # Star schema tables for BI (output)
│   ├── kpi_summary.xlsx                # Pre-aggregated business KPI tables (output)
│   └── data_quality_report.xlsx        # Data quality audit and metrics report (output)
└── README.md                           # Project documentation (this file)
```

## Pipeline Architecture

```text
[Raw Excel Source] 
       ↓
  [Ingestion]      (Multi-sheet parsing, skipping invalid header rows)
       ↓
  [Cleaning]       (Deduplication, standardisation, date parsing, price imputation)
       ↓
 [Enrichment]      (PII masking, joining product dimensions, calculating revenue)
       ↓
 [Aggregation]     (Generating pre-computed business KPIs and data quality metrics)
       ↓
   [Export]        (Saving clean CSV, Star Schema Excel, and KPI/DQ reports)
       ↓
  [Power BI]       (Importing Star Schema for interactive visualisations)
```

## How to Run

1. **Clone the repository:**
   ```bash
   git clone https://github.com/yourusername/neostats-retail-pipeline.git
   cd neostats-retail-pipeline
   ```
2. **Install dependencies:**
   ```bash
   pip install pandas openpyxl
   ```
3. **Place the source data:** Ensure the raw `USECASE_-_Data_Engineering.xlsx` file is located inside the `data/` directory.
4. **Execute the pipeline:**
   ```bash
   python Code/retail_pipeline.py
   ```
5. **Verify the outputs:** Check the console log for the pipeline summary. Output files will be generated in the `data/` directory.
6. **Open Power BI:** Follow the steps in `Power BI/dashboard_instructions.md` to load `powerbi_dataset.xlsx` into Power BI Desktop.

## Output Files

| File Name | Description | Location |
|---|---|---|
| `retail_final_cleaned.csv` | The master flattened fact table containing all cleaned and enriched transactions. | `data/` |
| `powerbi_dataset.xlsx` | A multi-sheet Star Schema workbook (Fact_Transactions, Dim_Product, Dim_Date, Dim_City) optimized for BI. | `data/` |
| `kpi_summary.xlsx` | 9 pre-computed pivot tables summarizing critical business metrics. | `data/` |
| `data_quality_report.xlsx` | An audit report tracking duplicates removed, imputed values, and parsing success rates. | `data/` |

## KPIs Delivered

1. **Total Revenue:** The sum total of generated revenue across all valid transactions.
2. **Total Transactions:** The absolute count of successful orders processed.
3. **Revenue by Category:** Total revenue grouped by canonical product categories (e.g., Electronics, Furniture).
4. **Revenue by City:** Geographical breakdown of sales performance.
5. **Revenue by Month:** Chronological revenue aggregation tracking seasonal growth.
6. **Top Products by Revenue:** The highest-performing individual SKUs.
7. **Revenue by Payment Method:** Distribution of revenue across Card, UPI, Netbanking, and Cash.
8. **Revenue by Channel:** Comparison of Online vs Offline sales figures.
9. **Avg Discount by Category:** The average promotional discount applied to distinct product groups.

## Data Quality Handling

The raw data contained multi-line headers, missing prices, invalid quantities, and inconsistent categorical data. The pipeline employs a robust cleaning strategy to standardize the dataset and ensure absolute analytical accuracy before BI consumption.

- **Dynamic parsing:** Handled varying multi-line headers and blank rows per Excel sheet dynamically.
- **Deduplication:** Removed exact duplicate rows and resolved payment status conflicts (successful transactions override failed ones).
- **Categorical normalization:** Standardized product names (Title Case, mapping variants to canonical names like "Laptop") and categories into 4 core groupings.
- **Date parsing:** Rescued dual date formats, converting both Excel integer serials and "MM-DD-YYYY" strings into a uniform `datetime` type.
- **Price imputation:** Imputed missing or zero prices in the transaction rows via a left join on the product dimension table.
- **Validation:** Filtered out invalid numeric data (e.g., negative or zero product quantities).

## PII Protection

To comply with data privacy standards, personally identifiable information (PII) was systematically masked before final export:
- **Email:** The `email` column was irreversibly hashed using **SHA-256** and stored as a secure hex string (`email_masked`). The original column was dropped.
- **Phone Number:** The `phone` column was obfuscated, leaving only the last 4 digits visible (e.g., `XXXXXX1234` stored as `phone_masked`). The original column was dropped.

## Assumptions

1. The provided Excel file `USECASE_-_Data_Engineering.xlsx` represents a standardized periodic data drop format.
2. The `Product Details` sheet contains the master source of truth for standard unit prices and canonical category mappings.
3. In the event of identical transaction IDs appearing across sheets, a row marked with a "successful" payment status takes precedence over a "failed" status.
4. Product quantities containing negative numbers or absolute zeros are invalid anomalies and are safely dropped.
5. Missing or zero transaction prices can be reliably imputed from the `Product Details` dimension table.
6. Transaction dates natively exist as either Excel serial integers or string arrays formatted as "MM-DD-YYYY".
7. The downstream Power BI dashboard is designed to natively ingest the flattened `powerbi_dataset.xlsx` using standard relational Star Schema models.

