# NeoStats Retail Data Engineering Pipeline
## Project Documentation

**Client:** ABC Retail Solutions
**Prepared by:** NeoStats Assessment Candidate
**Date:** June 2026
**Version:** 1.0

---

## 1. Project Overview

### Business Context

ABC Retail Solutions operates across multiple cities in India, selling products across four categories — Electronics, Clothing, Furniture, and Home Appliances — through both online and offline channels. The company collects transactional data from two separate retail systems (RETAIL DATA 1 and RETAIL DATA 2) which are periodically exported into Excel workbooks.

The business required a unified, clean, analysis-ready dataset to support revenue tracking, customer behaviour analysis, and product performance reporting through a Power BI dashboard.

### What Was Built

This project delivers a fully automated Python data engineering pipeline that:

- **Ingests** raw Excel data from three sheets: `PRODUCT DETAILS`, `RETAIL DATA 1`, and `RETAIL DATA 2`
- **Cleans** the data by resolving duplicates, invalid records, missing values, and inconsistent formatting
- **Standardises** categorical fields, product names, and date formats across both source systems
- **Protects** customer privacy by masking PII (email addresses and phone numbers)
- **Enriches** the transaction data with product dimension attributes
- **Computes** business-critical KPIs including total revenue, revenue by category, city, channel, and product
- **Exports** two analysis-ready output files: `retail_final_cleaned.csv` and `kpi_summary.xlsx`

---

## 2. Architecture Diagram

```
┌─────────────────────────────────────────────────────────────────────┐
│                        DATA ARCHITECTURE                            │
└─────────────────────────────────────────────────────────────────────┘

  ┌───────────────────────┐
  │   Excel Source File   │
  │  (USECASE_-_Data_     │
  │   Engineering.xlsx)   │
  │                       │
  │  ┌─────────────────┐  │
  │  │ PRODUCT DETAILS │  │
  │  ├─────────────────┤  │
  │  │  RETAIL DATA 1  │  │
  │  ├─────────────────┤  │
  │  │  RETAIL DATA 2  │  │
  └──┴─────────────────┴──┘
              │
              │  pd.read_excel()
              ▼
  ┌───────────────────────────────────────────────┐
  │            PYTHON PIPELINE                    │
  │           retail_pipeline.py                  │
  │                                               │
  │  Step 1  →  Data Ingestion                    │
  │  Step 2  →  Data Profiling                    │
  │  Step 3a →  Clean Transactions                │
  │  Step 3b →  Standardise Columns               │
  │  Step 3c →  Parse Dates                       │
  │  Step 3d →  Impute Missing Prices             │
  │  Step 3e →  Mask PII                          │
  │  Step 4  →  Concat + Enrich + Revenue         │
  │  Step 5  →  KPI Calculations                  │
  └───────────────────────────────────────────────┘
              │
       ┌──────┴──────┐
       │             │
       ▼             ▼
  ┌──────────┐  ┌──────────────┐
  │  CSV     │  │    Excel     │
  │ Output   │  │   Output     │
  │          │  │              │
  │ retail_  │  │ kpi_summary  │
  │ final_   │  │    .xlsx     │
  │ cleaned  │  │ (9 KPI tabs) │
  │  .csv    │  └──────────────┘
  └────┬─────┘
       │
       ▼
  ┌──────────────────┐
  │   Power BI       │
  │   Dashboard      │
  │                  │
  │ • Revenue KPIs   │
  │ • Category Split │
  │ • City Heatmap   │
  │ • Monthly Trend  │
  └──────────────────┘
```

---

## 3. Data Flow Description

### Step-by-Step Data Journey

1. **Ingestion**
   The pipeline reads three named sheets from the Excel workbook using `pd.read_excel()` with the `openpyxl` engine. Raw copies of the two retail DataFrames (`df_retail1_raw`, `df_retail2_raw`) are preserved before any modification to support auditability.

2. **Profiling**
   Each DataFrame is profiled: shape (rows × columns), column data types, null value counts per column, and a 3-row sample are printed to the console. This step runs before cleaning to surface issues early.

3. **Transaction Cleaning (3a)**
   Three sequential cleaning stages are applied to both retail DataFrames:
   - **Exact deduplication:** rows where every column matches are dropped (keep first)
   - **Payment conflict resolution:** where the same `transaction_id` has both a `"successful"` and `"failed"` payment record, the `"failed"` row is dropped
   - **Invalid quantity removal:** rows where `quantity ≤ 0` are dropped as they represent void, cancelled, or erroneous entries

4. **Column Standardisation (3b)**
   Raw categorical values are mapped to canonical forms using lookup dictionaries. String columns are normalised (title case / lowercase / strip) for consistent grouping downstream.

5. **Date Parsing (3c)**
   The `transaction_date` column contains two formats mixed together. Each cell is individually classified and parsed: Excel serial integers are converted using the Excel epoch (`1899-12-30 + N days`), and string dates are parsed with the `MM-DD-YYYY` format. Derived calendar columns (`transaction_year`, `transaction_month`, `transaction_month_name`, `transaction_quarter`) are appended.

6. **Price Imputation (3d)**
   Transactions with a missing or zero `price` are left-joined to the `PRODUCT DETAILS` table on `product_id`. The standard product price fills the gap. Only the price field is updated; no other product columns are retained in the transaction schema.

7. **PII Masking (3e)**
   Before any data is exported, customer identifiers are masked:
   - `email` → SHA-256 hex digest (irreversible one-way hash)
   - `phone` → `XXXXXX` prefix + last 4 digits only
   Both original columns are dropped and replaced with the masked variants.

8. **Final Dataset Assembly (Step 4)**
   Both cleaned retail DataFrames are tagged with a `data_source` column and concatenated. The combined DataFrame is enriched with `category_standard` and `standard_price` from the product dimension via a left join. Revenue is computed as `price × quantity × (1 − discount)`, rounded to 2 decimal places.

9. **KPI Export (Step 5)**
   Nine business KPI tables are computed via pandas `groupby` aggregations, printed to the console as a formatted summary report, and exported to `kpi_summary.xlsx` with one sheet per KPI.

---

## 4. Data Model

### Star Schema Design

The pipeline produces a logical star schema suitable for Power BI consumption.

```
                    ┌──────────────────────┐
                    │   DIMENSION TABLE    │
                    │   (PRODUCT DETAILS)  │
                    │                      │
                    │  product_id   (PK)   │
                    │  product_name        │
                    │  category            │
                    │  standard_price      │
                    └──────────┬───────────┘
                               │
                               │ 1 : Many
                               │
  ┌────────────────────────────▼─────────────────────────────────┐
  │                        FACT TABLE                            │
  │              (retail_final_cleaned.csv)                      │
  │                                                              │
  │  transaction_id          — unique transaction identifier     │
  │  product_id              — FK → PRODUCT DETAILS             │
  │  customer_id             — customer identifier               │
  │  quantity                — units purchased                   │
  │  price                   — unit price at time of purchase    │
  │  discount                — discount rate (0.0 – 1.0)        │
  │  revenue                 — price × qty × (1 − discount)     │
  │  transaction_date_parsed — clean datetime                    │
  │  transaction_year        — derived calendar dimension        │
  │  transaction_month       — derived calendar dimension        │
  │  transaction_month_name  — derived calendar dimension        │
  │  transaction_quarter     — derived calendar dimension        │
  │  purchase_location       — online / offline channel          │
  │  payment_method          — e.g. Credit Card, UPI, Cash      │
  │  payment_status          — successful / failed               │
  │  city                    — transaction city                  │
  │  category_standard       — enriched from product dim         │
  │  standard_price          — enriched from product dim         │
  │  email_masked            — SHA-256 hashed email              │
  │  phone_masked            — XXXXXX + last 4 digits            │
  │  data_source             — retail_data1 / retail_data2       │
  └──────────────────────────────────────────────────────────────┘
```

---

## 5. Data Quality Issues Found

The following issues were identified during profiling of the raw source data:

| # | Issue | Affected Column(s) | Source Sheet(s) |
|---|---|---|---|
| 1 | **Exact duplicate rows** — entire rows appearing more than once | All columns | RETAIL DATA 1, RETAIL DATA 2 |
| 2 | **Duplicate transaction IDs with conflicting payment status** — same `transaction_id` present with both `"successful"` and `"failed"` status | `transaction_id`, `payment_status` | Both retail sheets |
| 3 | **Invalid quantities** — rows with `quantity = 0`, `-1`, `-2` (voids, returns, data entry errors) | `quantity` | Both retail sheets |
| 4 | **Missing / zero prices** — some transactions had `price = 0` or `NaN`, making revenue calculation impossible | `price` | Both retail sheets |
| 5 | **Inconsistent category names** — same category encoded differently across rows and source systems (e.g. `"ELEC"`, `"elec"`, `"Electronics"`) | `category` | Both retail sheets |
| 6 | **Inconsistent product names** — product names with mixed casing, abbreviations, or extra characters (e.g. `"LAPTOP"`, `"laptop 15\""`, `"Lap Top"`) | `product_name` | Both retail sheets |
| 7 | **Mixed date formats** — `transaction_date` contained Excel serial integers (e.g. `45997`) alongside string dates in `MM-DD-YYYY` format (e.g. `"02-19-2026"`) | `transaction_date` | Both retail sheets |
| 8 | **PII in plain text** — customer email addresses and phone numbers stored as raw strings | `email`, `phone` | Both retail sheets |
| 9 | **Inconsistent string formatting** — leading/trailing whitespace, mixed case in `purchase_location`, `payment_method`, and `payment_status` | Multiple | Both retail sheets |

---

## 6. Cleaning Strategy

### 6.1 Exact Duplicate Removal

**Transformation:** `df.drop_duplicates(keep="first")`

**Business Justification:** Exact duplicate rows inflate transaction counts and revenue figures, leading to inaccurate KPIs. Retaining only the first occurrence preserves data integrity without discarding valid records.

### 6.2 Payment Conflict Resolution

**Transformation:** For any `transaction_id` that appears with both `"successful"` and `"failed"` status, the `"failed"` rows are dropped. Transaction IDs with only a single status are untouched.

**Business Justification:** A transaction that eventually succeeded should be counted once in revenue calculations. Retaining the failed duplicate would cause double-counting. A transaction that only has a `"failed"` record is retained — it may be relevant for failed-payment analysis.

### 6.3 Invalid Quantity Removal

**Transformation:** `df[df["quantity"] > 0]`

**Business Justification:** Quantities of zero or below represent voided, cancelled, or erroneously entered transactions. Including them in revenue aggregations would understate true sales figures.

### 6.4 Category Standardisation

**Transformation:** Lowercase the raw value, then apply a canonical mapping dictionary.

**Business Justification:** Inconsistent category names prevent correct grouping in Power BI slicers and pivot tables. A single canonical value per category ensures all transactions roll up correctly.

### 6.5 Product Name Standardisation

**Transformation:** Title-case the raw value, then apply substring keyword matching to map to canonical names.

**Business Justification:** Variant product name spellings fragment product-level revenue reporting, making top-product rankings unreliable.

### 6.6 Date Parsing

**Transformation:** Each cell is individually tested — numeric whole-numbers within the valid Excel serial range are converted via the Excel epoch; strings are parsed with `MM-DD-YYYY` format; a pandas fallback handles any remaining valid date strings.

**Business Justification:** A clean `datetime64` column is required to derive year, month, and quarter dimensions for trend analysis. Without parsing, time-series KPIs cannot be computed.

### 6.7 Missing Price Imputation

**Transformation:** Left-join transactions to `PRODUCT DETAILS` on `product_id`; fill `price` where null or zero with the standard product price.

**Business Justification:** Revenue cannot be calculated without a price. Using the product's standard price is a fair proxy for the transaction price in the absence of actual data, and is preferable to dropping the record entirely.

### 6.8 PII Masking

**Transformation:**
- `email` → SHA-256 hex digest (one-way, irreversible)
- `phone` → `XXXXXX` + last 4 digits

**Business Justification:** Customer PII must not be stored in plain text in analytical datasets. Hashing preserves referential integrity (the same email always produces the same hash) while making the raw value unrecoverable.

---

## 7. Transformation Logic

### 7.1 Category Mapping Table

| Raw Value(s) | Canonical Value |
|---|---|
| `ELEC`, `elec`, `electronics`, `Electronics` | `Electronics` |
| `CLOTH`, `cloth`, `clothing`, `Clothing` | `Clothing` |
| `FURN`, `furn`, `furniture`, `Furniture` | `Furniture` |
| `HOME`, `home`, `home appliances`, `Home Appliances` | `Home Appliances` |

### 7.2 Product Name Mapping

| Substring Match (case-insensitive) | Canonical Name |
|---|---|
| `laptop` | `Laptop` |
| `phone` | `Phone` |
| `shirt` | `Shirt` |
| `shoes` | `Shoes` |
| `dining table` | `Dining Table` |
| `mixer grinder` | `Mixer Grinder` |
| `refrigerator` | `Refrigerator` |
| `microwave` | `Microwave` |
| `sofa` | `Sofa` |
| `tv` | `TV` |

> **Note:** Multi-word keywords (e.g. `"dining table"`) are evaluated before single-word keywords to prevent partial matches.

### 7.3 Date Parsing Logic

```
For each cell in transaction_date:
    IF value is numeric AND 1 ≤ value ≤ 2,958,465:
        → Excel Serial: date = datetime(1899, 12, 30) + timedelta(days=int(value))
    ELSE IF value is string matching MM-DD-YYYY:
        → pd.to_datetime(value, format="%m-%d-%Y")
    ELSE:
        → pd.to_datetime(value, errors="coerce")   # fallback / NaT on failure
```

Derived columns produced:

| Column | Formula |
|---|---|
| `transaction_year` | `dt.year` |
| `transaction_month` | `dt.month` |
| `transaction_month_name` | `dt.strftime("%B")` |
| `transaction_quarter` | `"Q" + str(dt.quarter)` |

### 7.4 Revenue Formula

```
revenue = price × quantity × (1 − discount)
revenue = round(revenue, 2)
```

Where:
- `price` — unit price after missing-value imputation
- `quantity` — validated (> 0) units purchased
- `discount` — discount rate as a decimal (e.g. `0.10` = 10% off); treated as `0` if absent

---

## 8. KPI Definitions

### KPI 1 — Total Revenue
**Definition:** The sum of the `revenue` column across all successful transactions in `df_final`.
**Formula:** `SUM(revenue)`
**Purpose:** Top-line business health indicator. Reported in Indian number format (₹X,XX,XX,XXX.XX).

### KPI 2 — Revenue by Category
**Definition:** Total revenue grouped by product category (`Electronics`, `Clothing`, `Furniture`, `Home Appliances`).
**Formula:** `GROUP BY category → SUM(revenue) → ORDER BY revenue DESC`
**Purpose:** Identifies which product categories drive the most revenue, informing inventory and marketing decisions.

### KPI 3 — Revenue by City
**Definition:** Total revenue grouped by the city where the transaction took place.
**Formula:** `GROUP BY city → SUM(revenue) → ORDER BY revenue DESC`
**Purpose:** Reveals geographic revenue concentration to guide regional sales strategy and outlet expansion planning.

### KPI 4 — Revenue by Month
**Definition:** Total revenue grouped by year and calendar month, sorted chronologically.
**Formula:** `GROUP BY [transaction_year, transaction_month, transaction_month_name] → SUM(revenue)`
**Purpose:** Exposes seasonal trends, peak trading periods, and month-over-month growth patterns.

### KPI 5 — Top Products by Revenue
**Definition:** Total revenue per canonical product name, ranked descending.
**Formula:** `GROUP BY product_name → SUM(revenue) → ORDER BY revenue DESC`
**Purpose:** Identifies best-selling products to optimise stock allocation and promotional focus.

### KPI 6 — Revenue by Payment Method
**Definition:** Total revenue grouped by payment method (Credit Card, Debit Card, UPI, Cash, Net Banking, etc.).
**Formula:** `GROUP BY payment_method → SUM(revenue) → ORDER BY revenue DESC`
**Purpose:** Informs payment infrastructure investment and partnership decisions with payment providers.

### KPI 7 — Revenue by Channel (Online vs Offline)
**Definition:** Total revenue split by `purchase_location` (online vs in-store).
**Formula:** `GROUP BY purchase_location → SUM(revenue)`
**Purpose:** Measures the relative contribution of digital vs physical retail channels to guide omnichannel strategy.

### KPI 8 — Average Discount by Category
**Definition:** Mean discount rate per product category, expressed as a percentage.
**Formula:** `GROUP BY category → MEAN(discount) × 100`
**Purpose:** Monitors discount depth across categories to assess margin impact and promotional strategy effectiveness.

### KPI 9 — Total Orders by City
**Definition:** Count of distinct transaction records per city, ranked descending.
**Formula:** `GROUP BY city → COUNT(transaction_id) → ORDER BY count DESC`
**Purpose:** Measures transaction volume (not revenue) by location — useful for identifying high-footfall markets even where average order value may be lower.

---

## 9. Assumptions Made

1. **Failed payment transactions are excluded from revenue calculations.**
   Where a `transaction_id` exists with both `"successful"` and `"failed"` records, only the successful record is retained. Revenue is only recognised on completed transactions.

2. **The standard product price is a valid proxy for missing transaction prices.**
   When a transaction has a null or zero price, the corresponding product's standard price from `PRODUCT DETAILS` is used. This assumes the product was sold at its standard price in the absence of recorded data.

3. **Quantities of zero or below represent invalid or voided transactions.**
   Rows with `quantity ≤ 0` are excluded. It is assumed these do not represent valid sales events (e.g. they may be system errors, test entries, or cancelled orders with no return-processing offset).

4. **Excel serial dates and MM-DD-YYYY strings are the only two date formats in the source data.**
   The parser is designed to handle these two specific formats. Any other date format will fall back to pandas' generic parser; values that cannot be parsed will be stored as `NaT` and excluded from time-series KPIs.

5. **The discount column represents a rate (0.0–1.0), not a percentage (0–100).**
   Revenue is computed as `price × quantity × (1 − discount)`. A value of `0.10` means 10% off. If any source values are in percentage form, the revenue figure would be incorrect.

6. **Each `product_id` in the transaction data has a corresponding entry in `PRODUCT DETAILS`.**
   The enrichment join is a left join, so unmatched `product_id` values will produce null `category_standard` and `standard_price`. These records are retained in `df_final` but will not be included in category-level KPI groupings.

7. **Data from RETAIL DATA 1 and RETAIL DATA 2 represent the same business entity with compatible schemas.**
   Both sheets are concatenated into a single fact table. It is assumed the column names and value formats are consistent enough (after standardisation) to be unioned meaningfully.

---

## 10. Limitations & Future Improvements

### Current Limitations

- **No incremental load support:** The pipeline processes the full Excel file on every run. There is no mechanism to append only new records or detect changes from a previous run.

- **Static mapping dictionaries:** Category and product name mappings are hardcoded in the script. New product variants or category codes introduced in future data exports will not be automatically standardised — the mapping dictionaries must be manually updated.

- **Excel-only input:** The pipeline is built exclusively around the Excel workbook format. If the source system is migrated to a database, API, or CSV feed, the ingestion layer would need to be rewritten.

- **No data validation schema:** There is no formal schema validation (e.g. using `pandera` or `Great Expectations`) to assert column types, value ranges, or referential integrity before processing begins. Unexpected schema changes in the source file could cause silent errors downstream.

- **Single-threaded execution:** The pipeline runs sequentially. For very large datasets (millions of rows), processing time may become prohibitive without parallelisation.

- **No unit tests:** The cleaning and transformation functions are not covered by automated tests. Regressions from future changes cannot be automatically detected.

### Recommended Future Improvements

| Priority | Improvement | Benefit |
|---|---|---|
| High | Add `pandera` schema validation at ingestion | Catches data contract violations immediately |
| High | Write `pytest` unit tests for each transformation function | Prevents regressions, enables safe refactoring |
| High | Implement incremental / delta load logic | Avoids reprocessing historical data on each run |
| Medium | Externalise mapping dictionaries to a config file (JSON/YAML) | Business users can update mappings without touching code |
| Medium | Add a data quality report output (null rates, outlier flags, imputation counts) | Provides audit trail for downstream consumers |
| Medium | Migrate to a database source (PostgreSQL / BigQuery) | Enables production-scale data volumes |
| Low | Containerise with Docker | Ensures reproducible execution across environments |
| Low | Schedule via Apache Airflow or Azure Data Factory | Enables automated daily/weekly pipeline runs |

---

*End of Document*

---
> **Document version:** 1.0 | **Pipeline version:** 1.0.0 | **Last updated:** June 2026
