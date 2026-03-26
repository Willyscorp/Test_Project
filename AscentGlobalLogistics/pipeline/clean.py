"""
Data cleaning pipeline for Ascent Freight Pricing.
Each function is a single cleaning concern. The orchestrator calls them in order.
"""
import re
import logging

import pandas as pd
import numpy as np

logger = logging.getLogger(__name__)


def drop_header_rows(df):
    """Remove rows where column values are the header names (duplicate header rows)."""
    mask = df["shipment_id"] == "shipment_id"
    n = mask.sum()
    df = df[~mask].copy()
    logger.info(f"drop_header_rows: removed {n} rows")
    return df, n


def drop_null_rows(df):
    """Remove rows where all feature columns are null."""
    feature_cols = [c for c in df.columns if c != "shipment_id"]
    mask = df[feature_cols].isna().all(axis=1)
    n = mask.sum()
    df = df[~mask].copy()
    logger.info(f"drop_null_rows: removed {n} rows")
    return df, n


def standardize_equipment_type(df, mapping, drop_values):
    """Map equipment_type variants to canonical names; drop unmappable rows."""
    # Strip HTML tags first
    df["equipment_type"] = df["equipment_type"].apply(
        lambda x: re.sub(r"<[^>]+>", "", str(x)).strip() if pd.notna(x) else x
    )
    # Strip whitespace/non-breaking spaces
    df["equipment_type"] = df["equipment_type"].str.strip().str.replace("\xa0", "", regex=False)

    # Drop known bad values
    mask_drop = df["equipment_type"].isin(drop_values)
    n_dropped = mask_drop.sum()
    df = df[~mask_drop].copy()

    # Apply mapping
    df["equipment_type"] = df["equipment_type"].map(mapping)

    # Check for unmapped values
    unmapped = df["equipment_type"].isna() & df["equipment_type"].notna()
    n_unmapped = unmapped.sum()
    if n_unmapped > 0:
        logger.warning(f"standardize_equipment_type: {n_unmapped} unmapped values found")

    logger.info(f"standardize_equipment_type: dropped {n_dropped} rows, mapped to {df['equipment_type'].nunique()} categories")
    return df, n_dropped


def coerce_numeric_columns(df, columns):
    """Convert columns to numeric, turning garbage values into NaN."""
    for col in columns:
        before_nulls = df[col].isna().sum()
        df[col] = pd.to_numeric(df[col], errors="coerce")
        after_nulls = df[col].isna().sum()
        new_nulls = after_nulls - before_nulls
        if new_nulls > 0:
            logger.info(f"coerce_numeric_columns: {col} — {new_nulls} garbage values converted to NaN")
    # Also coerce target
    df["winning_bid_price"] = pd.to_numeric(df["winning_bid_price"], errors="coerce")
    return df


def remove_invalid_targets(df):
    """Drop rows where winning_bid_price is null, zero, or negative."""
    mask = (df["winning_bid_price"].isna()) | (df["winning_bid_price"] <= 0)
    n = mask.sum()
    df = df[~mask].copy()
    logger.info(f"remove_invalid_targets: removed {n} rows (null/zero/negative prices)")
    return df, n


def parse_dates(df):
    """Parse mixed date formats into a single datetime column."""
    df["shipment_date"] = pd.to_datetime(df["shipment_date"], format="mixed", dayfirst=False)
    n_null = df["shipment_date"].isna().sum()
    if n_null > 0:
        logger.info(f"parse_dates: {n_null} dates could not be parsed")
    return df


def cap_outliers(df, thresholds):
    """Cap values above the threshold for specified columns."""
    total_capped = 0
    for col, max_val in thresholds.items():
        if col in df.columns:
            mask = df[col] > max_val
            n = mask.sum()
            if n > 0:
                df.loc[mask, col] = max_val
                total_capped += n
                logger.info(f"cap_outliers: {col} — capped {n} values at {max_val}")
    # Also handle negative values in lead_time_hours
    if "lead_time_hours" in df.columns:
        neg_mask = df["lead_time_hours"] < 0
        n_neg = neg_mask.sum()
        if n_neg > 0:
            df.loc[neg_mask, "lead_time_hours"] = np.nan
            logger.info(f"cap_outliers: lead_time_hours — set {n_neg} negative values to NaN")
    return df, total_capped


def impute_missing(df, numeric_columns, strategy="median"):
    """Impute missing numeric values using the specified strategy."""
    impute_log = {}
    for col in numeric_columns:
        n_missing = df[col].isna().sum()
        if n_missing > 0:
            if strategy == "median":
                fill_val = df[col].median()
            else:
                fill_val = df[col].mean()
            df[col] = df[col].fillna(fill_val)
            impute_log[col] = {"missing": n_missing, "filled_with": round(fill_val, 2)}
            logger.info(f"impute_missing: {col} — filled {n_missing} NaNs with {strategy}={fill_val:.2f}")
    return df, impute_log


def run_cleaning_pipeline(df, config):
    """
    Orchestrator: runs all cleaning steps in order.
    Returns cleaned DataFrame and a log dict with row counts at each step.
    """
    log = {"initial_rows": len(df), "steps": []}

    def record(name, n_removed):
        log["steps"].append({"step": name, "rows_removed": n_removed, "rows_remaining": len(df)})

    df, n = drop_header_rows(df)
    record("drop_header_rows", n)

    df, n = drop_null_rows(df)
    record("drop_null_rows", n)

    df, n = standardize_equipment_type(df, config["equipment_mapping"], config["equipment_drop_values"])
    record("standardize_equipment_type", n)

    df = coerce_numeric_columns(df, config["numeric_columns"])
    record("coerce_numeric_columns", 0)

    df, n = remove_invalid_targets(df)
    record("remove_invalid_targets", n)

    df = parse_dates(df)
    record("parse_dates", 0)

    df, n_capped = cap_outliers(df, config["outlier_thresholds"])
    record("cap_outliers", 0)  # capping, not removing

    df, impute_log = impute_missing(df, config["numeric_columns"], config["imputation_strategy"])
    log["imputation"] = impute_log
    record("impute_missing", 0)

    log["final_rows"] = len(df)
    log["rows_removed_total"] = log["initial_rows"] - log["final_rows"]

    logger.info(f"Cleaning complete: {log['initial_rows']} -> {log['final_rows']} rows ({log['rows_removed_total']} removed)")
    return df, log
