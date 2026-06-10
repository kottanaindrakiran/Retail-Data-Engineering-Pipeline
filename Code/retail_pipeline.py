"""
=============================================================================
 Project   : NeoStats Retail Data Engineering Pipeline
 Client    : ABC Retail Solutions
 Author    : NeoStats Assessment Candidate
 Date      : June 2026
 Version   : 1.0.0
=============================================================================

Description
-----------
This script implements a production-grade, end-to-end retail data engineering
pipeline built for the NeoStats Assessment.  It ingests raw transactional
and product data from a multi-sheet Excel workbook, applies systematic data
quality transformations, enriches the dataset with product dimension
attributes, computes business KPIs, and exports analysis-ready outputs for
consumption by a Power BI dashboard.

Pipeline Steps
--------------
  Step 1  — Data Ingestion
              Loads three sheets (PRODUCT DETAILS, RETAIL DATA 1, RETAIL DATA 2)
              from the source Excel workbook into pandas DataFrames.

  Step 2  — Data Profiling
              Prints shape, dtypes, null counts, and sample rows for each
              DataFrame to support early data-quality awareness.

  Step 3  — Data Cleaning
      3a      Removes exact duplicate rows, resolves payment-status conflicts
              on transaction_id, and drops rows with invalid quantities.
      3b      Standardises category / product_name / payment columns to
              canonical values using lookup maps.
      3c      Parses the mixed-format transaction_date column (Excel serial
              numbers + MM-DD-YYYY strings) into proper datetime, and derives
              year / month / quarter dimension columns.
      3d      Imputes missing or zero prices from the product reference table
              via a left join on product_id.
      3e      Masks PII: SHA-256 hashes the email column, applies last-4-digit
              masking to the phone column.

  Step 4  — Build & Enrich Final Dataset
              Tags each source, concatenates both retail DataFrames, enriches
              with product dimension attributes (category_standard,
              standard_price), computes revenue = price × qty × (1 − discount),
              and exports retail_final_cleaned.csv.

  Step 5  — KPI Calculations
              Computes 9 business KPI tables (revenue by category, city, month,
              product, payment method, channel; avg discount; order counts),
              prints a formatted summary report, and exports kpi_summary.xlsx.

Outputs
-------
  data/retail_final_cleaned.csv   — cleaned, enriched transaction dataset
  data/kpi_summary.xlsx           — 9 KPI sheets + Summary cover sheet

Dependencies
------------
  pandas, openpyxl, numpy, hashlib, re, logging, datetime
  Install: pip install pandas openpyxl numpy

Usage
-----
  python Code/retail_pipeline.py
=============================================================================
"""

# ---------------------------------------------------------------------------
# IMPORTS
# ---------------------------------------------------------------------------

import os
import re
import hashlib
import logging
import numpy as np
import pandas as pd
import openpyxl                          # noqa: F401  (used by pandas engine)
from datetime import datetime

# ---------------------------------------------------------------------------
# LOGGING — INFO level, formatted output to console
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  [%(levelname)s]  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")

INPUT_FILE     = os.path.join(DATA_DIR, "USECASE - Data Engineering.xlsx")
OUTPUT_CLEAN   = os.path.join(DATA_DIR, "output_clean.xlsx")
OUTPUT_SUMMARY = os.path.join(DATA_DIR, "summary_report.xlsx")

# Sheet names in the workbook
SHEET_PRODUCT = "PRODUCT DETAILS"
SHEET_RETAIL1 = "RETAIL DATA 1"
SHEET_RETAIL2 = "RETAIL DATA 2"


# ---------------------------------------------------------------------------
# HELPER UTILITIES
# ---------------------------------------------------------------------------

def separator(title: str = "") -> None:
    """Print a visual section separator to stdout."""
    line = "=" * 70
    print(f"\n{line}")
    if title:
        print(f"  {title}")
        print(line)


# ---------------------------------------------------------------------------
# STEP 1 — DATA INGESTION
# ---------------------------------------------------------------------------

def ingest_data() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame,
                            pd.DataFrame, pd.DataFrame]:
    """
    Load the Excel workbook and read three named sheets into DataFrames.

    Returns
    -------
    df_product      : PRODUCT DETAILS sheet
    df_retail1      : RETAIL DATA 1 sheet
    df_retail2      : RETAIL DATA 2 sheet
    df_retail1_raw  : unmodified copy of df_retail1 (pre-cleaning snapshot)
    df_retail2_raw  : unmodified copy of df_retail2 (pre-cleaning snapshot)
    """
    separator("STEP 1 — DATA INGESTION")
    logger.info("Reading workbook: %s", INPUT_FILE)

    # ── Load all three sheets ──────────────────────────────────────────────
    # Each sheet has title/blank rows before the real column headers.
    # header= specifies the 0-indexed row that contains column names.
    #   PRODUCT DETAILS : row 3  (blank, title, blank, headers)
    #   RETAIL DATA 1   : row 2  (title, blank, headers)
    #   RETAIL DATA 2   : row 4  (blank, blank, title, blank, headers)
    df_product = pd.read_excel(
        INPUT_FILE,
        sheet_name=SHEET_PRODUCT,
        engine="openpyxl",
        header=3,
    )

    df_retail1 = pd.read_excel(
        INPUT_FILE,
        sheet_name=SHEET_RETAIL1,
        engine="openpyxl",
        header=2,
    )

    df_retail2 = pd.read_excel(
        INPUT_FILE,
        sheet_name=SHEET_RETAIL2,
        engine="openpyxl",
        header=4,
    )

    # ── Save raw (pre-cleaning) copies before any transformation ──────────
    df_retail1_raw = df_retail1.copy()
    df_retail2_raw = df_retail2.copy()

    # ── Print shape and column names for each DataFrame ───────────────────
    for label, df in [
        (SHEET_PRODUCT, df_product),
        (SHEET_RETAIL1, df_retail1),
        (SHEET_RETAIL2, df_retail2),
    ]:
        print(f"\n  [{label}]")
        print(f"    Shape   : {df.shape}  ({df.shape[0]} rows × {df.shape[1]} cols)")
        print(f"    Columns : {list(df.columns)}")

    # ── Structured completion log ──────────────────────────────────────────
    logger.info(
        "Data ingestion complete. "
        "retail_data1: %d rows, retail_data2: %d rows, product_details: %d rows",
        len(df_retail1),
        len(df_retail2),
        len(df_product),
    )

    return df_product, df_retail1, df_retail2, df_retail1_raw, df_retail2_raw


# ---------------------------------------------------------------------------
# STEP 2 — DATA PROFILING  (placeholder)
# ---------------------------------------------------------------------------

def profile_data(df_product: pd.DataFrame,
                 df_retail1: pd.DataFrame,
                 df_retail2: pd.DataFrame) -> None:
    """
    Print a profiling summary for each DataFrame:
    dtypes, null counts, and a top-3 sample.
    """
    separator("STEP 2 — DATA PROFILING")

    for label, df in [
        (SHEET_PRODUCT, df_product),
        (SHEET_RETAIL1, df_retail1),
        (SHEET_RETAIL2, df_retail2),
    ]:
        print(f"\n--- {label} ---")
        print(f"  Shape      : {df.shape}")
        print(f"  Dtypes:\n{df.dtypes.to_string()}")
        print(f"\n  Null counts:\n{df.isnull().sum().to_string()}")
        print(f"\n  Sample (top 3 rows):\n{df.head(3).to_string(index=False)}")

    logger.info("Data profiling complete.")


# ---------------------------------------------------------------------------
# STEP 3 — DATA CLEANING
# ---------------------------------------------------------------------------

def clean_transactions(df: pd.DataFrame, source_name: str) -> pd.DataFrame:
    """
    Clean a retail transactions DataFrame in three sequential stages:

    Stage 1 — Exact duplicate removal
        Drops rows where every column is identical; keeps the first occurrence.

    Stage 2 — Transaction-ID conflict resolution
        Where multiple rows share the same transaction_id but differ in
        payment_status, keeps only rows where payment_status == 'successful'
        and drops 'failed' duplicates.  Rows with any other status combination
        (e.g. two 'successful' rows) are left untouched by this stage.

    Stage 3 — Invalid quantity removal
        Drops any row where the quantity column is <= 0 (covers 0, negative
        integers, and negative floats).

    Parameters
    ----------
    df          : Input transactions DataFrame.
    source_name : Human-readable label used in log messages (e.g. 'RETAIL DATA 1').

    Returns
    -------
    pd.DataFrame : Cleaned copy of the input DataFrame.
    """
    separator(f"STEP 3 — CLEAN TRANSACTIONS: {source_name}")
    df = df.copy()                          # never mutate the caller's object

    # ------------------------------------------------------------------
    # Stage 1: Remove exact duplicate rows (all columns match)
    # ------------------------------------------------------------------
    rows_before = len(df)
    df = df.drop_duplicates(keep="first")
    exact_dupes_removed = rows_before - len(df)
    logger.info(
        "[%s] Stage 1 — Exact duplicates removed : %d rows  (%d → %d)",
        source_name, exact_dupes_removed, rows_before, len(df),
    )

    # ------------------------------------------------------------------
    # Stage 2: Resolve transaction_id conflicts on payment_status
    #   • Find transaction_ids that appear with BOTH 'successful' AND
    #     'failed' payment_status values.
    #   • For those IDs, drop the 'failed' rows only.
    #   • Rows whose IDs have a single status (or only 'failed') are
    #     untouched by this stage.
    # ------------------------------------------------------------------
    rows_before = len(df)

    # Detect column names (case-insensitive)
    tx_col  = next((c for c in df.columns if "transaction" in c.lower() and "id" in c.lower()), None)
    pay_col = next((c for c in df.columns if "payment" in c.lower() and "status" in c.lower()), None)

    if tx_col and pay_col:
        # Normalise payment_status to lowercase for reliable comparison
        df[pay_col] = df[pay_col].astype(str).str.strip().str.lower()

        # Identify IDs that have at least one 'successful' row
        ids_with_success = set(
            df.loc[df[pay_col] == "successful", tx_col].unique()
        )

        # Drop rows where: ID has a 'successful' counterpart AND this row is 'failed'
        conflict_mask = (
            df[tx_col].isin(ids_with_success) & (df[pay_col] == "failed")
        )
        df = df[~conflict_mask].reset_index(drop=True)

        conflict_rows_removed = rows_before - len(df)
        logger.info(
            "[%s] Stage 2 — Payment conflict rows removed: %d rows  "
            "(kept 'successful', dropped 'failed' for %d transaction_id(s))",
            source_name,
            conflict_rows_removed,
            len(ids_with_success & set(df.loc[df[pay_col] == "successful", tx_col])),
        )
    else:
        logger.warning(
            "[%s] Stage 2 — Skipped: could not locate 'transaction_id' "
            "and/or 'payment_status' columns. Found columns: %s",
            source_name, list(df.columns),
        )
        conflict_rows_removed = 0

    # ------------------------------------------------------------------
    # Stage 3: Remove rows with invalid quantity (<= 0)
    # ------------------------------------------------------------------
    rows_before = len(df)

    qty_col = next(
        (c for c in df.columns if any(k in c.lower() for k in ["quantity", "qty", "units"])),
        None,
    )

    if qty_col:
        df = df[df[qty_col] > 0].reset_index(drop=True)
        invalid_qty_removed = rows_before - len(df)
        logger.info(
            "[%s] Stage 3 — Invalid quantity rows removed : %d rows  "
            "(quantity <= 0 in column '%s')",
            source_name, invalid_qty_removed, qty_col,
        )
    else:
        logger.warning(
            "[%s] Stage 3 — Skipped: could not locate a quantity column. "
            "Found columns: %s",
            source_name, list(df.columns),
        )
        invalid_qty_removed = 0

    # ------------------------------------------------------------------
    # Summary
    # ------------------------------------------------------------------
    total_removed = exact_dupes_removed + conflict_rows_removed + invalid_qty_removed
    logger.info(
        "[%s] Cleaning complete — total rows removed: %d  |  final row count: %d",
        source_name, total_removed, len(df),
    )
    return df


def clean_dataframe(df: pd.DataFrame, label: str) -> pd.DataFrame:
    """
    Generic cleaner used for non-transactional sheets (e.g. PRODUCT DETAILS):
      - Remove exact duplicates
      - Strip / normalise string columns
      - Parse date columns
      - Impute missing values
      - Correct negative numeric values
    """
    separator(f"STEP 3 — DATA CLEANING: {label}")
    original_rows = len(df)

    # 1. Remove exact duplicates
    df = df.drop_duplicates()
    logger.info("  Dropped %d duplicate rows.", original_rows - len(df))

    # 2. Strip whitespace & normalise strings
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip()
        if any(kw in col.lower() for kw in ["category", "region", "status", "type", "name"]):
            df[col] = df[col].str.title()

    # 3. Parse date columns
    for col in df.columns:
        if any(kw in col.lower() for kw in ["date", "time", "day"]):
            try:
                df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")
                logger.info("  Parsed '%s' as datetime.", col)
            except Exception:
                pass

    # 4. Impute missing values
    for col in df.columns:
        null_count = int(df[col].isnull().sum())
        if null_count == 0:
            continue
        if pd.api.types.is_numeric_dtype(df[col]):
            median_val = df[col].median()
            df[col] = df[col].fillna(median_val)
            logger.info("  Filled %d nulls in '%s' with median (%.2f).", null_count, col, median_val)
        elif pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].ffill()
            logger.info("  Forward-filled %d nulls in datetime column '%s'.", null_count, col)
        else:
            df[col] = df[col].fillna("Unknown")
            logger.info("  Filled %d nulls in '%s' with 'Unknown'.", null_count, col)

    # 5. Enforce non-negative numerics
    for col in df.select_dtypes(include="number").columns:
        neg_count = int((df[col] < 0).sum())
        if neg_count > 0:
            df[col] = df[col].abs()
            logger.info("  Corrected %d negative values in '%s'.", neg_count, col)

    logger.info("  Cleaning done. Final shape: %s", df.shape)
    return df


# ---------------------------------------------------------------------------
# STEP 3b — COLUMN STANDARDISATION
# ---------------------------------------------------------------------------

# Canonical category mapping — keys normalised to lowercase+stripped at runtime
_CATEGORY_MAP: dict[str, str] = {
    "elec"            : "Electronics",
    "electronics"     : "Electronics",
    "cloth"           : "Clothing",
    "clothing"        : "Clothing",
    "furn"            : "Furniture",
    "furniture"       : "Furniture",
    "home"            : "Home Appliances",
    "home appliances" : "Home Appliances",
}

# Canonical product_name mapping
# Each key is a lowercase substring; first match wins.
_PRODUCT_KEYWORD_MAP: list[tuple[str, str]] = [
    ("laptop"        , "Laptop"),
    ("phone"         , "Phone"),
    ("shirt"         , "Shirt"),
    ("shoes"         , "Shoes"),
    ("dining table"  , "Dining Table"),   # must come before "tv" / "sofa" etc.
    ("mixer grinder" , "Mixer Grinder"),
    ("refrigerator"  , "Refrigerator"),
    ("microwave"     , "Microwave"),
    ("sofa"          , "Sofa"),
    ("tv"            , "TV"),
]


def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Standardise categorical and textual columns in a retail transactions DataFrame.

    Column transformations applied (columns detected case-insensitively):

    category
        Maps raw variations to one of four canonical values:
        Electronics | Clothing | Furniture | Home Appliances

    product_name
        Title-cases the raw value, then maps common product name variants
        to a single canonical form (e.g. 'LAPTOP', 'Laptop 15"' → 'Laptop').

    purchase_location
        Lowercase + strip whitespace.

    payment_method
        Title case.

    payment_status
        Lowercase.

    Parameters
    ----------
    df : Input DataFrame (modified in-place on a copy).

    Returns
    -------
    pd.DataFrame : Standardised copy of the input DataFrame.
    """
    df = df.copy()

    # ── Helper: find first matching column (case-insensitive keyword search) ──
    def _find_col(keyword: str) -> str | None:
        return next((c for c in df.columns if keyword in c.lower()), None)

    # ------------------------------------------------------------------
    # 1. category
    # ------------------------------------------------------------------
    cat_col = _find_col("category")
    if cat_col:
        before_vals = sorted(df[cat_col].dropna().unique().tolist())
        logger.info("  [category] unique values BEFORE standardisation: %s", before_vals)

        df[cat_col] = (
            df[cat_col]
            .astype(str)
            .str.strip()
            .str.lower()
            .map(lambda v: _CATEGORY_MAP.get(v, v))   # unmapped values pass through
        )

        after_vals = sorted(df[cat_col].dropna().unique().tolist())
        logger.info("  [category] unique values AFTER  standardisation: %s", after_vals)
    else:
        logger.warning("  standardize_columns: 'category' column not found. Columns: %s", list(df.columns))

    # ------------------------------------------------------------------
    # 2. product_name
    # ------------------------------------------------------------------
    prod_col = next(
        (c for c in df.columns if "product" in c.lower() and "name" in c.lower()),
        None,
    )
    if prod_col:
        # Title-case first, then apply keyword-based canonical mapping
        df[prod_col] = df[prod_col].astype(str).str.strip().str.title()

        def _map_product(val: str) -> str:
            lower = val.lower()
            for keyword, canonical in _PRODUCT_KEYWORD_MAP:
                if keyword in lower:
                    return canonical
            return val   # no match → keep title-cased original

        df[prod_col] = df[prod_col].map(_map_product)
        logger.info(
            "  [product_name] unique values after standardisation: %s",
            sorted(df[prod_col].unique().tolist()),
        )
    else:
        logger.warning("  standardize_columns: 'product_name' column not found.")

    # ------------------------------------------------------------------
    # 3. purchase_location → lowercase + strip
    # ------------------------------------------------------------------
    loc_col = _find_col("location")
    if loc_col:
        df[loc_col] = df[loc_col].astype(str).str.strip().str.lower()
        logger.info(
            "  [purchase_location] normalised to lowercase. Unique: %s",
            sorted(df[loc_col].unique().tolist()),
        )
    else:
        logger.warning("  standardize_columns: 'purchase_location' column not found.")

    # ------------------------------------------------------------------
    # 4. payment_method → title case
    # ------------------------------------------------------------------
    method_col = next(
        (c for c in df.columns
         if "payment" in c.lower() and "method" in c.lower()), None
    )
    if method_col:
        df[method_col] = df[method_col].astype(str).str.strip().str.title()
        logger.info(
            "  [payment_method] normalised to title case. Unique: %s",
            sorted(df[method_col].unique().tolist()),
        )
    else:
        logger.warning("  standardize_columns: 'payment_method' column not found.")

    # ------------------------------------------------------------------
    # 5. payment_status → lowercase
    # ------------------------------------------------------------------
    status_col = next(
        (c for c in df.columns
         if "payment" in c.lower() and "status" in c.lower()), None
    )
    if status_col:
        df[status_col] = df[status_col].astype(str).str.strip().str.lower()
        logger.info(
            "  [payment_status] normalised to lowercase. Unique: %s",
            sorted(df[status_col].unique().tolist()),
        )
    else:
        logger.warning("  standardize_columns: 'payment_status' column not found.")

    logger.info("  Column standardisation complete.")
    return df


# ---------------------------------------------------------------------------
# STEP 3c — DATE PARSING
# ---------------------------------------------------------------------------

def _parse_single_date(value) -> "pd.Timestamp | pd.NaT":
    """
    Parse one cell from the transaction_date column.

    Supports two raw formats:
      • Excel serial integer  (e.g. 45997)
          → datetime(1899, 12, 30) + timedelta(days=int(value))
      • String in MM-DD-YYYY  (e.g. "02-19-2026")
          → pd.to_datetime(value, format="%m-%d-%Y")

    Returns pd.NaT if the value cannot be parsed.
    """
    if pd.isna(value):
        return pd.NaT

    # ── Branch 1: Excel serial number ──────────────────────────────────────
    # Integers and floats that look like whole numbers are treated as serials.
    try:
        numeric = float(value)
        if numeric == int(numeric) and 1 <= int(numeric) <= 2_958_465:  # 1900-01-01 to 9999-12-31
            return pd.Timestamp(
                datetime(1899, 12, 30) + __import__("datetime").timedelta(days=int(numeric))
            )
    except (ValueError, TypeError):
        pass

    # ── Branch 2: String date MM-DD-YYYY ───────────────────────────────────
    try:
        return pd.to_datetime(str(value).strip(), format="%m-%d-%Y")
    except (ValueError, TypeError):
        pass

    # ── Fallback: let pandas try anything reasonable ────────────────────────
    try:
        return pd.to_datetime(str(value).strip(), dayfirst=False, errors="coerce")
    except Exception:
        return pd.NaT


def parse_dates(df: pd.DataFrame) -> pd.DataFrame:
    """
    Parse the mixed-format 'transaction_date' column into a clean datetime
    column and derive calendar dimension columns from it.

    Input column handled
    --------------------
    transaction_date
        May contain a mix of:
          • Excel serial integers  (e.g. 45997)
          • Strings in MM-DD-YYYY  (e.g. "02-19-2026")

    New columns created
    -------------------
    transaction_date_parsed  : datetime64[ns]
    transaction_year         : int   (e.g. 2026)
    transaction_month        : int   (1 – 12)
    transaction_month_name   : str   (e.g. "January")
    transaction_quarter      : str   (e.g. "Q1")

    Counts of serial-parsed vs string-parsed vs failed cells are logged
    at INFO level to surface data provenance issues early.

    Returns
    -------
    pd.DataFrame with the new columns appended.
    """
    df = df.copy()

    # Locate the transaction_date column (case-insensitive)
    date_col = next(
        (c for c in df.columns if "transaction" in c.lower() and "date" in c.lower()),
        None,
    )

    if date_col is None:
        logger.warning(
            "parse_dates: 'transaction_date' column not found. "
            "Found columns: %s", list(df.columns)
        )
        return df

    logger.info("parse_dates: processing column '%s' (%d cells)", date_col, len(df))

    # ── Per-cell classification counters ────────────────────────────────────
    serial_count  = 0
    string_count  = 0
    failed_count  = 0
    parsed_dates  = []

    for raw in df[date_col]:
        if pd.isna(raw):
            parsed_dates.append(pd.NaT)
            failed_count += 1
            continue

        # Try Excel serial first
        serial_parsed = False
        try:
            numeric = float(raw)
            if numeric == int(numeric) and 1 <= int(numeric) <= 2_958_465:
                from datetime import timedelta
                parsed_dates.append(
                    pd.Timestamp(datetime(1899, 12, 30) + timedelta(days=int(numeric)))
                )
                serial_count += 1
                serial_parsed = True
        except (ValueError, TypeError):
            pass

        if serial_parsed:
            continue

        # Try MM-DD-YYYY string
        try:
            parsed_dates.append(
                pd.to_datetime(str(raw).strip(), format="%m-%d-%Y")
            )
            string_count += 1
            continue
        except (ValueError, TypeError):
            pass

        # Fallback — let pandas infer
        fallback = pd.to_datetime(str(raw).strip(), errors="coerce")
        parsed_dates.append(fallback)
        if pd.isna(fallback):
            failed_count += 1
            logger.debug("parse_dates: could not parse value %r", raw)
        else:
            string_count += 1   # count fallback successes as string-parsed

    # ── Assign parsed column ────────────────────────────────────────────────
    df["transaction_date_parsed"] = pd.to_datetime(parsed_dates)

    logger.info(
        "parse_dates: parsed %d serial | %d string | %d failed  "
        "(total %d cells)",
        serial_count, string_count, failed_count, len(df),
    )

    # ── Derive calendar dimension columns ───────────────────────────────────
    ts = df["transaction_date_parsed"]

    df["transaction_year"]       = ts.dt.year.astype("Int64")         # nullable int
    df["transaction_month"]      = ts.dt.month.astype("Int64")
    df["transaction_month_name"] = ts.dt.strftime("%B").where(ts.notna(), other=pd.NA)
    df["transaction_quarter"]    = ts.dt.quarter.map(
        lambda q: f"Q{int(q)}" if pd.notna(q) else pd.NA
    )

    logger.info(
        "parse_dates: derived columns → "
        "transaction_year, transaction_month, transaction_month_name, transaction_quarter"
    )
    logger.info(
        "parse_dates: date range → %s  to  %s",
        ts.min().date() if ts.notna().any() else "N/A",
        ts.max().date() if ts.notna().any() else "N/A",
    )

    return df


# ---------------------------------------------------------------------------
# STEP 3d — MISSING PRICE IMPUTATION
# ---------------------------------------------------------------------------

def fix_missing_prices(df: pd.DataFrame, df_product: pd.DataFrame) -> pd.DataFrame:
    """
    Impute missing or zero prices in a transaction DataFrame using the
    standard product price from the PRODUCT DETAILS reference table.

    Algorithm
    ---------
    1. Detect the join key column (contains "product" and "id") and the
       price column (contains "price") in both DataFrames — case-insensitive.
    2. Left-join the transaction DataFrame to a slim product lookup table
       (product_id + standard_price only) to avoid polluting the schema.
    3. Where the transaction price is NaN or 0, replace it with the
       standard price from the product table.
    4. Drop the temporary merge column; restore the original column order.
    5. Log the number of prices imputed.

    Parameters
    ----------
    df         : Cleaned transaction DataFrame (RETAIL DATA 1 or 2).
    df_product : PRODUCT DETAILS DataFrame — must contain product_id and price.

    Returns
    -------
    pd.DataFrame : Transactions with corrected price column; schema unchanged.
    """
    df = df.copy()

    # ── 1. Locate join-key and price columns (case-insensitive) ────────────
    tx_id_col = next(
        (c for c in df.columns
         if "product" in c.lower() and "id" in c.lower()), None
    )
    tx_price_col = next(
        (c for c in df.columns if "price" in c.lower()), None
    )

    prod_id_col = next(
        (c for c in df_product.columns
         if "product" in c.lower() and "id" in c.lower()), None
    )
    prod_price_col = next(
        (c for c in df_product.columns if "price" in c.lower()), None
    )

    # Guard: skip gracefully if any required column is missing
    if not all([tx_id_col, tx_price_col, prod_id_col, prod_price_col]):
        logger.warning(
            "fix_missing_prices: required columns not found. "
            "tx_id='%s', tx_price='%s', prod_id='%s', prod_price='%s' — skipping.",
            tx_id_col, tx_price_col, prod_id_col, prod_price_col,
        )
        return df

    logger.info(
        "fix_missing_prices: join '%s' → '%s'  |  price cols: '%s' ← '%s'",
        tx_id_col, prod_id_col, tx_price_col, prod_price_col,
    )

    # ── 2. Build slim product lookup (id + price only) ─────────────────────
    # Rename product price to a temp name to avoid collision on merge
    _TEMP_COL = "__std_price__"
    product_lookup = (
        df_product[[prod_id_col, prod_price_col]]
        .drop_duplicates(subset=[prod_id_col])
        .rename(columns={prod_id_col: tx_id_col, prod_price_col: _TEMP_COL})
    )

    # ── 3. Left join ────────────────────────────────────────────────────────
    # Cast join key to string on both sides to avoid dtype mismatch (object vs int64)
    original_cols = list(df.columns)          # snapshot before merge
    df[tx_id_col] = df[tx_id_col].astype(str)
    product_lookup[tx_id_col] = product_lookup[tx_id_col].astype(str)
    df = df.merge(product_lookup, on=tx_id_col, how="left")

    # ── 4. Identify rows that need imputation (null or zero price) ──────────
    needs_impute = df[tx_price_col].isna() | (df[tx_price_col] == 0)
    impute_count = int(needs_impute.sum())

    # Fill only where needed
    df.loc[needs_impute, tx_price_col] = df.loc[needs_impute, _TEMP_COL]

    # ── 5. Drop the temp column; restore original column order ──────────────
    df = df.drop(columns=[_TEMP_COL])
    # Preserve original column order (merge may have shifted columns)
    df = df[original_cols]

    logger.info(
        "fix_missing_prices: %d price(s) imputed from product reference table "
        "(out of %d transactions).",
        impute_count, len(df),
    )

    # Warn if any prices are still missing after imputation
    still_null = int(df[tx_price_col].isna().sum())
    if still_null:
        logger.warning(
            "fix_missing_prices: %d transaction(s) still have no price after "
            "imputation (product_id not matched in reference table).",
            still_null,
        )

    return df


# ---------------------------------------------------------------------------
# STEP 3e — PII MASKING
# ---------------------------------------------------------------------------

def mask_pii(df: pd.DataFrame) -> pd.DataFrame:
    """
    Mask personally identifiable information (PII) before the DataFrame is
    written to any output file.

    Transformations applied
    -----------------------
    email
        Each value is SHA-256 hashed (UTF-8 encoded) and stored as a 64-char
        hex digest in a new column ``email_masked``.  The original ``email``
        column is then dropped.  NaN / null values produce the string
        "<null>" (hashed) so the column remains fully populated.

    phone
        All characters except the last 4 are replaced with "XXXXXX", e.g.
        "9876543210" → "XXXXXX3210".  The result is stored in a new column
        ``phone_masked`` and the original ``phone`` column is dropped.
        Non-string values are coerced to str before masking.  Values with
        fewer than 4 characters are stored as "XXXXXX" (fully masked).

    Column detection is case-insensitive (keyword match on column name).
    If a column is not found a WARNING is logged and the function continues
    processing the remaining columns.

    Parameters
    ----------
    df : Transaction DataFrame to mask.

    Returns
    -------
    pd.DataFrame : Copy of ``df`` with PII columns replaced by masked variants.
    """
    df = df.copy()
    n = len(df)

    # ── Helper ─────────────────────────────────────────────────────────────────
    def _find_col(keyword: str) -> str | None:
        return next((c for c in df.columns if keyword in c.lower()), None)

    # ------------------------------------------------------------------
    # 1. Email — SHA-256 hash
    # ------------------------------------------------------------------
    email_col = _find_col("email")

    if email_col:
        def _hash_email(val) -> str:
            raw = "<null>" if pd.isna(val) else str(val).strip().lower()
            return hashlib.sha256(raw.encode("utf-8")).hexdigest()

        df["email_masked"] = df[email_col].map(_hash_email)
        df = df.drop(columns=[email_col])
        logger.info(
            "  [email] SHA-256 hashed → 'email_masked'. Original '%s' dropped.",
            email_col,
        )
    else:
        logger.warning("  mask_pii: 'email' column not found — skipped. Columns: %s", list(df.columns))

    # ------------------------------------------------------------------
    # 2. Phone — XXXXXX + last 4 digits
    # ------------------------------------------------------------------
    phone_col = _find_col("phone")

    if phone_col:
        def _mask_phone(val) -> str:
            if pd.isna(val):
                return "XXXXXX"
            s = re.sub(r"\D", "", str(val))   # strip non-digits
            if len(s) <= 4:
                return "XXXXXX"               # too short — mask entirely
            return f"XXXXXX{s[-4:]}"

        df["phone_masked"] = df[phone_col].map(_mask_phone)
        df = df.drop(columns=[phone_col])
        logger.info(
            "  [phone] last-4 masking applied → 'phone_masked'. Original '%s' dropped.",
            phone_col,
        )
    else:
        logger.warning("  mask_pii: 'phone' column not found — skipped. Columns: %s", list(df.columns))

    # ------------------------------------------------------------------
    # Summary log (exact wording as specified)
    # ------------------------------------------------------------------
    logger.info("PII masking complete. %d records masked.", n)

    return df


# ---------------------------------------------------------------------------
# STEP 4 — BUILD & ENRICH FINAL DATASET
# ---------------------------------------------------------------------------

def build_final_dataset(
    df_retail1: pd.DataFrame,
    df_retail2: pd.DataFrame,
    df_product: pd.DataFrame,
) -> pd.DataFrame:
    """
    Combine both cleaned retail DataFrames into a single analysis-ready
    dataset, enrich it with product dimension attributes, and compute revenue.

    Steps
    -----
    1. Tag each DataFrame with a ``data_source`` identifier column.
    2. Concatenate into ``df_final`` and reset the index.
    3. Left-join ``df_product`` to add ``category_standard`` and
       ``standard_price`` from the product dimension table.
    4. Compute ``revenue = price * quantity * (1 - discount)``, rounded
       to 2 decimal places.  Column names are detected case-insensitively.
    5. Print ``df_final.shape``, ``df_final.dtypes``, and ``df_final.head(5)``.
    6. Export ``df_final`` to ``data/retail_final_cleaned.csv`` (index=False).
    7. Log the final row and column counts.

    Returns
    -------
    pd.DataFrame : The fully enriched, analysis-ready final DataFrame.
    """
    separator("STEP 4 — BUILD & ENRICH FINAL DATASET")

    # ------------------------------------------------------------------
    # 1. Tag data source
    # ------------------------------------------------------------------
    df_retail1 = df_retail1.copy()
    df_retail2 = df_retail2.copy()

    df_retail1["data_source"] = "retail_data1"
    df_retail2["data_source"] = "retail_data2"
    logger.info("  Tagged df_retail1 as 'retail_data1', df_retail2 as 'retail_data2'.")

    # ------------------------------------------------------------------
    # 2. Concatenate and reset index
    # ------------------------------------------------------------------
    df_final = pd.concat(
        [df_retail1, df_retail2],
        ignore_index=True,
        sort=False,
    )
    logger.info(
        "  Concatenated DataFrames: %d rows × %d cols",
        df_final.shape[0], df_final.shape[1],
    )

    # ------------------------------------------------------------------
    # 3. Enrich with product dimension — left join
    # ------------------------------------------------------------------
    # Detect join key (case-insensitive)
    tx_id_col   = next((c for c in df_final.columns
                        if "product" in c.lower() and "id" in c.lower()), None)
    prod_id_col = next((c for c in df_product.columns
                        if "product" in c.lower() and "id" in c.lower()), None)

    # Detect which product columns to pull in
    prod_cat_col   = next((c for c in df_product.columns
                           if "category" in c.lower()), None)
    prod_price_col = next((c for c in df_product.columns
                           if "price" in c.lower()), None)

    if tx_id_col and prod_id_col and (prod_cat_col or prod_price_col):
        # Build slim dimension lookup — only columns we need
        dim_cols = [prod_id_col]
        rename_map: dict[str, str] = {prod_id_col: tx_id_col}
        if prod_cat_col:
            dim_cols.append(prod_cat_col)
            rename_map[prod_cat_col] = "category_standard"
        if prod_price_col:
            dim_cols.append(prod_price_col)
            rename_map[prod_price_col] = "standard_price"

        product_dim = (
            df_product[dim_cols]
            .drop_duplicates(subset=[prod_id_col])
            .rename(columns=rename_map)
        )

        # Cast join key to str on both sides to prevent dtype mismatch
        df_final[tx_id_col] = df_final[tx_id_col].astype(str)
        product_dim[tx_id_col] = product_dim[tx_id_col].astype(str)
        df_final = df_final.merge(product_dim, on=tx_id_col, how="left")
        new_cols = [v for v in rename_map.values() if v != tx_id_col]
        logger.info(
            "  Enriched with product dimension columns: %s  (%d unmatched rows)",
            new_cols,
            int(df_final[new_cols[0]].isna().sum()) if new_cols else 0,
        )
    else:
        logger.warning(
            "  Skipped product enrichment — could not locate required columns. "
            "tx_id='%s', prod_id='%s'", tx_id_col, prod_id_col,
        )

    # ------------------------------------------------------------------
    # 4. Calculate revenue = price * quantity * (1 - discount)
    # ------------------------------------------------------------------
    price_col    = next((c for c in df_final.columns
                         if c.lower() == "price"
                         or ("price" in c.lower() and "standard" not in c.lower())), None)
    qty_col      = next((c for c in df_final.columns
                         if any(k in c.lower() for k in ["quantity", "qty", "units"])), None)
    discount_col = next((c for c in df_final.columns
                         if "discount" in c.lower()), None)

    if price_col and qty_col:
        p = pd.to_numeric(df_final[price_col],    errors="coerce").fillna(0)
        q = pd.to_numeric(df_final[qty_col],      errors="coerce").fillna(0)

        if discount_col:
            d = pd.to_numeric(df_final[discount_col], errors="coerce").fillna(0)
            df_final["revenue"] = (p * q * (1 - d)).round(2)
            logger.info(
                "  Computed 'revenue' = %s × %s × (1 - %s)  [rounded to 2 dp]",
                price_col, qty_col, discount_col,
            )
        else:
            df_final["revenue"] = (p * q).round(2)
            logger.info(
                "  Computed 'revenue' = %s × %s  (no discount column found)",
                price_col, qty_col,
            )
    else:
        logger.warning(
            "  Revenue not computed — could not locate price ('%s') or "
            "quantity ('%s') column.", price_col, qty_col,
        )

    # ------------------------------------------------------------------
    # 5. Print shape, dtypes, head
    # ------------------------------------------------------------------
    separator("df_final — SHAPE / DTYPES / SAMPLE")
    print(f"\n  Shape   : {df_final.shape}  ({df_final.shape[0]:,} rows × {df_final.shape[1]} cols)")
    print(f"\n  Dtypes:\n{df_final.dtypes.to_string()}")
    print(f"\n  Head (5 rows):\n{df_final.head(5).to_string(index=False)}")

    # ------------------------------------------------------------------
    # 6. Export to CSV
    # ------------------------------------------------------------------
    output_csv = os.path.join(DATA_DIR, "retail_final_cleaned.csv")
    df_final.to_csv(output_csv, index=False)
    logger.info(
        "Final dataset exported: %d rows, %d columns  →  %s",
        df_final.shape[0], df_final.shape[1], output_csv,
    )

    return df_final


# =============================================================================
# === KPI CALCULATIONS ===
# =============================================================================

def _fmt_inr(value: float) -> str:
    """
    Format a number in Indian numbering system (lakhs / crores).
    e.g.  12345678.90  →  "1,23,45,678.90"
    """
    # Split integer and decimal parts
    s = f"{abs(value):.2f}"
    integer_part, decimal_part = s.split(".")

    # Apply Indian grouping: last 3 digits, then groups of 2
    n = len(integer_part)
    if n <= 3:
        grouped = integer_part
    else:
        # First group: rightmost 3 digits
        grouped = integer_part[-3:]
        remaining = integer_part[:-3]
        # Subsequent groups: 2 digits each
        while remaining:
            grouped = remaining[-2:] + "," + grouped
            remaining = remaining[:-2]

    sign = "-" if value < 0 else ""
    return f"{sign}Rs.{grouped}.{decimal_part}"


def compute_kpis(df_final: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """
    Compute all business KPI tables from the enriched final DataFrame
    and export them to a single Excel workbook (one sheet per KPI).

    KPIs computed
    -------------
    1.  total_revenue            — scalar sum (printed in Indian format)
    2.  revenue_by_category      — grouped sum, descending
    3.  revenue_by_city          — grouped sum, descending
    4.  revenue_by_month         — grouped sum, sorted by year then month
    5.  top_products_by_revenue  — grouped sum, descending
    6.  revenue_by_payment_method— grouped sum, descending
    7.  revenue_by_channel       — grouped sum (online vs offline)
    8.  avg_discount_by_category — grouped mean
    9.  total_orders_by_city     — transaction_id count, descending

    Output
    ------
    Exports ``data/kpi_summary.xlsx`` with each KPI on its own sheet.
    Prints a formatted summary report to the console.

    Returns
    -------
    dict[str, pd.DataFrame] : All KPI tables keyed by name.
    """
    separator("STEP 5 — KPI CALCULATIONS")

    kpis: dict[str, pd.DataFrame] = {}

    # Helper — resolve column name case-insensitively, with fallback
    def _col(keyword: str, df: pd.DataFrame = df_final) -> str | None:
        return next((c for c in df.columns if keyword in c.lower()), None)

    rev_col      = _col("revenue")
    cat_col      = _col("category")
    city_col     = _col("city")
    yr_col       = "transaction_year"
    mo_col       = "transaction_month"
    mo_name_col  = "transaction_month_name"
    prod_col     = _col("product_name") or _col("product")
    method_col   = next((c for c in df_final.columns
                         if "payment" in c.lower() and "method" in c.lower()), None)
    loc_col      = _col("location") or _col("channel")
    disc_col     = _col("discount")
    tx_col       = next((c for c in df_final.columns
                         if "transaction" in c.lower() and "id" in c.lower()), None)

    # ------------------------------------------------------------------
    # KPI 1 — Total Revenue
    # ------------------------------------------------------------------
    total_revenue = float(df_final[rev_col].sum()) if rev_col else 0.0

    # ------------------------------------------------------------------
    # KPI 2 — Revenue by Category
    # ------------------------------------------------------------------
    if rev_col and cat_col:
        kpis["Revenue_by_Category"] = (
            df_final.groupby(cat_col)[rev_col]
            .sum()
            .sort_values(ascending=False)
            .reset_index()
            .rename(columns={cat_col: "Category", rev_col: "Revenue"})
        )

    # ------------------------------------------------------------------
    # KPI 3 — Revenue by City
    # ------------------------------------------------------------------
    if rev_col and city_col:
        kpis["Revenue_by_City"] = (
            df_final.groupby(city_col)[rev_col]
            .sum()
            .sort_values(ascending=False)
            .reset_index()
            .rename(columns={city_col: "City", rev_col: "Revenue"})
        )

    # ------------------------------------------------------------------
    # KPI 4 — Revenue by Month
    # ------------------------------------------------------------------
    month_dims = [c for c in [yr_col, mo_col, mo_name_col] if c in df_final.columns]
    if rev_col and len(month_dims) >= 2:
        kpis["Revenue_by_Month"] = (
            df_final.groupby(month_dims)[rev_col]
            .sum()
            .reset_index()
            .sort_values([yr_col, mo_col])
            .rename(columns={rev_col: "Revenue"})
        )

    # ------------------------------------------------------------------
    # KPI 5 — Top Products by Revenue
    # ------------------------------------------------------------------
    if rev_col and prod_col:
        kpis["Top_Products_by_Revenue"] = (
            df_final.groupby(prod_col)[rev_col]
            .sum()
            .sort_values(ascending=False)
            .reset_index()
            .rename(columns={prod_col: "Product", rev_col: "Revenue"})
        )

    # ------------------------------------------------------------------
    # KPI 6 — Revenue by Payment Method
    # ------------------------------------------------------------------
    if rev_col and method_col:
        kpis["Revenue_by_Payment_Method"] = (
            df_final.groupby(method_col)[rev_col]
            .sum()
            .sort_values(ascending=False)
            .reset_index()
            .rename(columns={method_col: "Payment_Method", rev_col: "Revenue"})
        )

    # ------------------------------------------------------------------
    # KPI 7 — Revenue by Channel (online vs offline)
    # ------------------------------------------------------------------
    if rev_col and loc_col:
        kpis["Revenue_by_Channel"] = (
            df_final.groupby(loc_col)[rev_col]
            .sum()
            .reset_index()
            .rename(columns={loc_col: "Channel", rev_col: "Revenue"})
        )

    # ------------------------------------------------------------------
    # KPI 8 — Average Discount by Category
    # ------------------------------------------------------------------
    if disc_col and cat_col:
        kpis["Avg_Discount_by_Category"] = (
            df_final.groupby(cat_col)[disc_col]
            .mean()
            .reset_index()
            .rename(columns={cat_col: "Category", disc_col: "Avg_Discount"})
        )
        kpis["Avg_Discount_by_Category"]["Avg_Discount_Pct"] = (
            kpis["Avg_Discount_by_Category"]["Avg_Discount"].mul(100).round(2)
        )

    # ------------------------------------------------------------------
    # KPI 9 — Total Orders by City
    # ------------------------------------------------------------------
    if city_col and tx_col:
        kpis["Total_Orders_by_City"] = (
            df_final.groupby(city_col)[tx_col]
            .count()
            .sort_values(ascending=False)
            .reset_index()
            .rename(columns={city_col: "City", tx_col: "Total_Orders"})
        )

    # ------------------------------------------------------------------
    # Console summary report
    # ------------------------------------------------------------------
    separator("KPI SUMMARY REPORT")

    print(f"\n  {'='*60}")
    print(f"  {'NeoStats — ABC Retail Solutions  |  KPI Summary':^60}")
    print(f"  {'='*60}")
    print(f"\n  > Total Revenue            :  {_fmt_inr(total_revenue)}")
    print(f"  > Total Transactions       :  {len(df_final):,}")
    print(f"  > Total Products           :  {df_final[prod_col].nunique() if prod_col else 'N/A'}")
    print(f"  > Total Cities             :  {df_final[city_col].nunique() if city_col else 'N/A'}")

    if "Revenue_by_Category" in kpis:
        print(f"\n  > Revenue by Category:")
        for _, row in kpis["Revenue_by_Category"].iterrows():
            print(f"      {row['Category']:<22}  {_fmt_inr(row['Revenue'])}")

    if "Top_Products_by_Revenue" in kpis:
        top5 = kpis["Top_Products_by_Revenue"].head(5)
        print(f"\n  > Top 5 Products by Revenue:")
        for i, (_, row) in enumerate(top5.iterrows(), 1):
            print(f"      {i}. {row['Product']:<20}  {_fmt_inr(row['Revenue'])}")

    if "Revenue_by_Payment_Method" in kpis:
        print(f"\n  > Revenue by Payment Method:")
        for _, row in kpis["Revenue_by_Payment_Method"].iterrows():
            print(f"      {str(row['Payment_Method']):<22}  {_fmt_inr(row['Revenue'])}")

    if "Revenue_by_Channel" in kpis:
        print(f"\n  > Revenue by Channel (Online vs Offline):")
        for _, row in kpis["Revenue_by_Channel"].iterrows():
            print(f"      {str(row['Channel']):<22}  {_fmt_inr(row['Revenue'])}")

    if "Avg_Discount_by_Category" in kpis:
        print(f"\n  > Avg Discount by Category:")
        for _, row in kpis["Avg_Discount_by_Category"].iterrows():
            print(f"      {row['Category']:<22}  {row['Avg_Discount_Pct']:.2f}%")

    if "Total_Orders_by_City" in kpis:
        top_city = kpis["Total_Orders_by_City"].iloc[0]
        print(f"\n  > Top City by Orders       :  {top_city['City']} ({top_city['Total_Orders']:,} orders)")

    if "Revenue_by_Month" in kpis:
        peak = kpis["Revenue_by_Month"].loc[
            kpis["Revenue_by_Month"]["Revenue"].idxmax()
        ]
        month_name = peak.get(mo_name_col, str(int(peak[mo_col])))
        print(f"  > Peak Revenue Month       :  {month_name} {int(peak[yr_col])}  "
              f"({_fmt_inr(peak['Revenue'])})")

    print(f"\n  {'='*60}\n")

    # ------------------------------------------------------------------
    # Export to Excel  — one sheet per KPI
    # ------------------------------------------------------------------
    kpi_output = os.path.join(DATA_DIR, "kpi_summary.xlsx")
    logger.info("Exporting %d KPI tables → %s", len(kpis), kpi_output)

    with pd.ExcelWriter(kpi_output, engine="openpyxl") as writer:
        # Write a cover sheet with the scalar KPIs
        cover = pd.DataFrame({
            "KPI"   : ["Total Revenue", "Total Transactions", "Total Products", "Total Cities"],
            "Value" : [
                _fmt_inr(total_revenue),
                f"{len(df_final):,}",
                str(df_final[prod_col].nunique()) if prod_col else "N/A",
                str(df_final[city_col].nunique()) if city_col else "N/A",
            ],
        })
        cover.to_excel(writer, sheet_name="Summary", index=False)

        for sheet_name, kpi_df in kpis.items():
            kpi_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            logger.info("  Written KPI sheet: '%s' (%d rows)", sheet_name[:31], len(kpi_df))

    logger.info("KPI export complete → %s", kpi_output)
    return kpis


def export_for_powerbi(df_final: pd.DataFrame, df_product: pd.DataFrame) -> None:
    """
    Export the final dataset into a multi-sheet Excel workbook structured
    specifically for Power BI ingestion (Star Schema format).
    """
    output_path = os.path.join(DATA_DIR, "powerbi_dataset.xlsx")
    
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Sheet 1: Fact_Transactions
        fact_cols = [
            "transaction_id", "customer_id", "customer_name", "product_id", "product_name", 
            "category", "category_standard", "purchase_location", "city", 
            "transaction_date_parsed", "transaction_year", "transaction_month", 
            "transaction_month_name", "transaction_quarter", "quantity", "price", 
            "standard_price", "discount", "payment_method", "data_source", "revenue"
        ]
        fact_cols = [c for c in fact_cols if c in df_final.columns]
        df_fact = df_final[fact_cols].copy()
        df_fact = df_fact.rename(columns={
            "transaction_date_parsed": "Date",
            "purchase_location": "Channel",
            "category_standard": "Category"
        })
        if "Date" in df_fact.columns:
            df_fact = df_fact.sort_values(by="Date", ascending=True)
        df_fact.to_excel(writer, sheet_name="Fact_Transactions", index=False)
        
        # Sheet 2: Dim_Product
        df_dim_prod = df_product[["product_id", "product_name", "category", "price"]].copy()
        df_dim_prod = df_dim_prod.rename(columns={"price": "Standard_Price"})
        df_dim_prod = df_dim_prod.drop_duplicates()
        df_dim_prod.to_excel(writer, sheet_name="Dim_Product", index=False)
        
        # Sheet 3: Dim_Date
        if "Date" in df_fact.columns:
            min_date = df_fact["Date"].min()
            max_date = df_fact["Date"].max()
            date_range = pd.date_range(start=min_date, end=max_date)
            df_dim_date = pd.DataFrame({"Date": date_range})
            df_dim_date["Year"] = df_dim_date["Date"].dt.year
            df_dim_date["Month"] = df_dim_date["Date"].dt.month
            df_dim_date["Month_Name"] = df_dim_date["Date"].dt.strftime("%B")
            df_dim_date["Quarter"] = "Q" + df_dim_date["Date"].dt.quarter.astype(str)
            df_dim_date["Week_of_Year"] = df_dim_date["Date"].dt.isocalendar().week
            df_dim_date["Day_of_Week"] = df_dim_date["Date"].dt.dayofweek + 1
            df_dim_date["Day_Name"] = df_dim_date["Date"].dt.day_name()
            df_dim_date["Is_Weekend"] = df_dim_date["Date"].dt.dayofweek >= 5
            df_dim_date.to_excel(writer, sheet_name="Dim_Date", index=False)
            
        # Sheet 4: Dim_City
        if "city" in df_fact.columns:
            cities = sorted(df_fact["city"].dropna().unique().tolist())
            df_dim_city = pd.DataFrame({
                "City_ID": range(1, len(cities) + 1),
                "City": cities
            })
            df_dim_city.to_excel(writer, sheet_name="Dim_City", index=False)
            
        # Sheet 5: KPI_Summary
        total_rev = df_final["revenue"].sum()
        total_txn = len(df_final)
        total_cust = df_final["customer_id"].nunique()
        avg_rev = total_rev / total_txn if total_txn > 0 else 0
        
        top_cat = df_final.groupby("category_standard")["revenue"].sum().idxmax() if "category_standard" in df_final.columns else "N/A"
        top_city = df_final.groupby("city")["revenue"].sum().idxmax() if "city" in df_final.columns else "N/A"
        top_prod = df_final.groupby("product_name")["revenue"].sum().idxmax() if "product_name" in df_final.columns else "N/A"
        
        df_kpi = pd.DataFrame({
            "KPI_Name": [
                "Total Revenue", 
                "Total Transactions", 
                "Total Unique Customers", 
                "Avg Revenue per Transaction", 
                "Top Category by Revenue", 
                "Top City by Revenue", 
                "Top Product by Revenue"
            ],
            "KPI_Value": [
                _fmt_inr(total_rev),
                f"{total_txn:,}",
                f"{total_cust:,}",
                _fmt_inr(avg_rev),
                str(top_cat),
                str(top_city),
                str(top_prod)
            ]
        })
        df_kpi.to_excel(writer, sheet_name="KPI_Summary", index=False)
        
    logger.info("Power BI dataset exported: %s", output_path)


def generate_data_quality_report(df_retail1_raw: pd.DataFrame, df_retail2_raw: pd.DataFrame, df_final: pd.DataFrame) -> None:
    """
    Generate a Data Quality Report Excel file comparing raw and clean datasets.
    """
    output_path = os.path.join(DATA_DIR, "data_quality_report.xlsx")
    
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        
        # --- Sheet 1: Raw_vs_Clean ---
        def get_metrics(df_raw):
            exact_dups = len(df_raw) - len(df_raw.drop_duplicates())
            failed_payment = len(df_raw[df_raw["payment_status"].astype(str).str.lower() == "failed"]) if "payment_status" in df_raw.columns else 0
            invalid_qty = len(df_raw[pd.to_numeric(df_raw["quantity"], errors="coerce").fillna(0) <= 0]) if "quantity" in df_raw.columns else 0
            missing_price = len(df_raw[pd.to_numeric(df_raw["price"], errors="coerce").isna() | (pd.to_numeric(df_raw["price"], errors="coerce") == 0)]) if "price" in df_raw.columns else 0
            null_dates = df_raw["transaction_date"].isna().sum() if "transaction_date" in df_raw.columns else 0
            uniq_tx = df_raw["transaction_id"].nunique() if "transaction_id" in df_raw.columns else 0
            uniq_cust = df_raw["customer_id"].nunique() if "customer_id" in df_raw.columns else 0
            uniq_prod = df_raw["product_id"].nunique() if "product_id" in df_raw.columns else 0
            return {
                "Total Rows": len(df_raw),
                "Exact Duplicate Rows Removed": exact_dups,
                "Failed Payment Rows Removed": failed_payment,
                "Invalid Quantity Rows Removed": invalid_qty,
                "Missing Prices Imputed": missing_price,
                "Rows with Null Dates": null_dates,
                "Unique Transaction IDs": uniq_tx,
                "Unique Customers": uniq_cust,
                "Unique Products": uniq_prod,
            }
            
        m1 = get_metrics(df_retail1_raw)
        m2 = get_metrics(df_retail2_raw)
        
        m_final = {
            "Total Rows": len(df_final),
            "Exact Duplicate Rows Removed": "N/A",
            "Failed Payment Rows Removed": "N/A",
            "Invalid Quantity Rows Removed": "N/A",
            "Missing Prices Imputed": "N/A",
            "Rows with Null Dates": df_final["transaction_date_parsed"].isna().sum() if "transaction_date_parsed" in df_final.columns else 0,
            "Unique Transaction IDs": df_final["transaction_id"].nunique() if "transaction_id" in df_final.columns else 0,
            "Unique Customers": df_final["customer_id"].nunique() if "customer_id" in df_final.columns else 0,
            "Unique Products": df_final["product_id"].nunique() if "product_id" in df_final.columns else 0,
        }
        
        df_raw_vs_clean = pd.DataFrame({
            "Metric": list(m1.keys()),
            "Retail_Data1_Raw": list(m1.values()),
            "Retail_Data2_Raw": list(m2.values()),
            "Final_Combined": list(m_final.values())
        })
        df_raw_vs_clean.to_excel(writer, sheet_name="Raw_vs_Clean", index=False)
        
        # --- Sheet 2: Category_Standardization ---
        raw_cats = pd.concat([df_retail1_raw["category"], df_retail2_raw["category"]]).dropna().unique()
        df_cat_std = pd.DataFrame({"Raw_Category": raw_cats})
        df_cat_std["Standardized_Category"] = df_cat_std["Raw_Category"].astype(str).str.strip().str.lower().map(lambda v: _CATEGORY_MAP.get(v, v))
        df_cat_std.to_excel(writer, sheet_name="Category_Standardization", index=False)
        
        # --- Sheet 3: Date_Parse_Summary ---
        raw_dates = pd.concat([df_retail1_raw["transaction_date"], df_retail2_raw["transaction_date"]]).dropna()
        serial_mask = pd.to_numeric(raw_dates, errors="coerce").notna()
        str_mask = ~serial_mask
        
        excel_serial_count = serial_mask.sum()
        string_count = str_mask.sum()
        
        failed_parses = df_final["transaction_date_parsed"].isna().sum() if "transaction_date_parsed" in df_final.columns else 0
        min_date = df_final["transaction_date_parsed"].min().strftime("%Y-%m-%d") if not df_final["transaction_date_parsed"].isna().all() else "N/A"
        max_date = df_final["transaction_date_parsed"].max().strftime("%Y-%m-%d") if not df_final["transaction_date_parsed"].isna().all() else "N/A"
        
        df_date_parse = pd.DataFrame({
            "Metric": [
                "Dates parsed from Excel serial format",
                "Dates parsed from MM-DD-YYYY string format",
                "Failed date parses (NaT)",
                "Date range (min to max)"
            ],
            "Value": [
                int(excel_serial_count),
                int(string_count),
                int(failed_parses),
                f"{min_date} to {max_date}"
            ]
        })
        df_date_parse.to_excel(writer, sheet_name="Date_Parse_Summary", index=False)
        
        # --- Sheet 4: Null_Analysis ---
        null_counts = df_final.isna().sum()
        null_pct = (null_counts / len(df_final)) * 100
        
        df_nulls = pd.DataFrame({
            "Column_Name": null_counts.index,
            "Null_Count": null_counts.values,
            "Null_Percentage": null_pct.values.round(2)
        })
        df_nulls.to_excel(writer, sheet_name="Null_Analysis", index=False)
        
    logger.info("Data quality report exported: %s", output_path)


# ---------------------------------------------------------------------------
# MAIN ENTRYPOINT
# ---------------------------------------------------------------------------

def main() -> None:
    """
    Orchestrate the full NeoStats Retail Data Engineering Pipeline.

    Calls each processing step in sequence with clearly printed section
    headers, handles file-not-found errors gracefully, and prints a
    completion summary with output file paths at the end.

    Steps executed
    --------------
    1  Data Ingestion       — load Excel sheets into DataFrames
    2  Data Profiling       — print shape / dtypes / nulls / samples
    3  Data Cleaning        — deduplicate, standardise, parse dates,
                              impute prices, mask PII
    4  Build Final Dataset  — concat, enrich, compute revenue, export CSV
    5  KPI Calculations     — 9 KPI tables, console report, Excel export
    """
    start_time = datetime.now()

    print("\n" + "=" * 70)
    print("  NeoStats Retail Data Engineering Pipeline")
    print("  Client : ABC Retail Solutions  |  Started:", start_time.strftime("%Y-%m-%d %H:%M:%S"))
    print("=" * 70)

    # ================================================================
    # === STEP 1: DATA INGESTION ===
    # ================================================================
    print("\n=== STEP 1: DATA INGESTION ===")
    try:
        df_product, df_retail1, df_retail2, df_retail1_raw, df_retail2_raw = ingest_data()
    except FileNotFoundError:
        logger.error(
            "Source file not found: %s\n"
            "  → Please place 'USECASE_-_Data_Engineering.xlsx' in the data/ folder "
            "and re-run the pipeline.",
            INPUT_FILE,
        )
        raise SystemExit(1)
    except Exception as exc:
        logger.error("Unexpected error during data ingestion: %s", exc, exc_info=True)
        raise SystemExit(1)

    # ================================================================
    # === STEP 2: DATA PROFILING ===
    # ================================================================
    print("\n=== STEP 2: DATA PROFILING ===")
    profile_data(df_product, df_retail1, df_retail2)

    # ================================================================
    # === STEP 3: DATA CLEANING ===
    # ================================================================
    print("\n=== STEP 3: DATA CLEANING ===")

    # 3a — Transaction cleaning (dedup / payment conflicts / invalid qty)
    print("\n--- Step 3a: Transaction Cleaning ---")
    df_product = clean_dataframe(df_product, SHEET_PRODUCT)

    print(f"\n  [RETAIL DATA 1]  Rows before cleaning : {len(df_retail1):,}")
    df_retail1 = clean_transactions(df_retail1, SHEET_RETAIL1)
    print(f"  [RETAIL DATA 1]  Rows after  cleaning : {len(df_retail1):,}")

    print(f"\n  [RETAIL DATA 2]  Rows before cleaning : {len(df_retail2):,}")
    df_retail2 = clean_transactions(df_retail2, SHEET_RETAIL2)
    print(f"  [RETAIL DATA 2]  Rows after  cleaning : {len(df_retail2):,}")

    # 3b — Column standardisation
    print("\n--- Step 3b: Column Standardisation ---")
    df_retail1 = standardize_columns(df_retail1)
    df_retail2 = standardize_columns(df_retail2)

    # 3c — Date parsing
    print("\n--- Step 3c: Date Parsing ---")
    df_retail1 = parse_dates(df_retail1)
    df_retail2 = parse_dates(df_retail2)

    # 3d — Missing price imputation
    print("\n--- Step 3d: Missing Price Imputation ---")
    df_retail1 = fix_missing_prices(df_retail1, df_product)
    df_retail2 = fix_missing_prices(df_retail2, df_product)

    # 3e — PII masking
    print("\n--- Step 3e: PII Masking ---")
    df_retail1 = mask_pii(df_retail1)
    df_retail2 = mask_pii(df_retail2)

    # ================================================================
    # === STEP 4: BUILD & ENRICH FINAL DATASET ===
    # ================================================================
    print("\n=== STEP 4: BUILD & ENRICH FINAL DATASET ===")
    df_final = build_final_dataset(df_retail1, df_retail2, df_product)

    # ================================================================
    # === STEP 5: KPI CALCULATIONS ===
    # ================================================================
    print("\n=== STEP 5: KPI CALCULATIONS ===")
    compute_kpis(df_final)

    # ================================================================
    # === STEP 6: POWER BI EXPORT ===
    # ================================================================
    print("\n=== STEP 6: POWER BI EXPORT ===")
    export_for_powerbi(df_final, df_product)

    # ================================================================
    # === STEP 7: DATA QUALITY REPORT ===
    # ================================================================
    print("\n=== STEP 7: DATA QUALITY REPORT ===")
    generate_data_quality_report(df_retail1_raw, df_retail2_raw, df_final)

    # ================================================================
    # === PIPELINE COMPLETE ===
    # ================================================================
    elapsed = datetime.now() - start_time
    print("\n" + "=" * 70)
    print("  PIPELINE COMPLETE [OK]")
    print(f"  Elapsed time : {elapsed.seconds // 60}m {elapsed.seconds % 60}s")
    print(f"  Final rows   : {df_final.shape[0]:,}")
    print(f"  Final cols   : {df_final.shape[1]}")
    print("  Output files :")
    print(f"    -> {os.path.join(DATA_DIR, 'retail_final_cleaned.csv')}")
    print(f"    -> {os.path.join(DATA_DIR, 'kpi_summary.xlsx')}")
    print(f"    -> {os.path.join(DATA_DIR, 'powerbi_dataset.xlsx')}")
    print(f"    -> {os.path.join(DATA_DIR, 'data_quality_report.xlsx')}")
    print("=" * 70 + "\n")

    logger.info(
        "Pipeline finished in %ds. Final dataset: %d rows × %d cols.",
        elapsed.seconds, df_final.shape[0], df_final.shape[1],
    )


if __name__ == "__main__":
    main()
