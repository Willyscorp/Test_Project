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


def remove_negative_values(df, columns):
    """Remove rows with negative values in physical columns (weight, volume, bids)."""
    total_removed = 0
    for col in columns:
        if col in df.columns:
            mask = df[col] < 0
            n = mask.sum()
            if n > 0:
                df = df[~mask].copy()
                total_removed += n
                logger.info(f"remove_negative_values: {col} — removed {n} rows with negative values")
    return df, total_removed


def remove_invalid_dates(df, min_date="2023-01-01", max_date="2026-12-31"):
    """Remove rows with dates outside a reasonable range."""
    df["shipment_date"] = pd.to_datetime(df["shipment_date"], format="mixed", dayfirst=False, errors="coerce")
    before = len(df)
    mask = (df["shipment_date"] < min_date) | (df["shipment_date"] > max_date) | df["shipment_date"].isna()
    n = mask.sum()
    df = df[~mask].copy()
    logger.info(f"remove_invalid_dates: removed {n} rows outside [{min_date}, {max_date}]")
    return df, n


def remove_outliers_iqr(df, columns, multiplier=1.5):
    """Remove rows with outliers using IQR method. Returns df and log of what was removed."""
    total_removed = 0
    outlier_log = {}
    for col in columns:
        if col not in df.columns:
            continue
        q1 = df[col].quantile(0.25)
        q3 = df[col].quantile(0.75)
        iqr = q3 - q1
        lower = q1 - multiplier * iqr
        upper = q3 + multiplier * iqr
        mask = (df[col] < lower) | (df[col] > upper)
        n = mask.sum()
        if n > 0:
            df = df[~mask].copy()
            total_removed += n
            outlier_log[col] = {
                "Q1": round(q1, 2), "Q3": round(q3, 2), "IQR": round(iqr, 2),
                "lower_bound": round(lower, 2), "upper_bound": round(upper, 2),
                "removed": n,
            }
            logger.info(f"remove_outliers_iqr: {col} — removed {n} rows outside [{lower:.2f}, {upper:.2f}]")
    return df, total_removed, outlier_log


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


def run_strict_cleaning_pipeline(df, config):
    """
    Strict orchestrator: removes outliers and all garbage data instead of capping.
    Returns cleaned DataFrame and a detailed log dict.
    """
    log = {"initial_rows": len(df), "steps": []}

    def record(name, n_removed):
        log["steps"].append({"step": name, "rows_removed": n_removed, "rows_remaining": len(df)})

    # Step 1: Drop header rows
    df, n = drop_header_rows(df)
    record("drop_header_rows", n)

    # Step 2: Drop fully null rows
    df, n = drop_null_rows(df)
    record("drop_null_rows", n)

    # Step 3: Standardize equipment type
    df, n = standardize_equipment_type(df, config["equipment_mapping"], config["equipment_drop_values"])
    record("standardize_equipment_type", n)

    # Step 4: Coerce numeric columns
    df = coerce_numeric_columns(df, config["numeric_columns"])
    record("coerce_numeric_columns", 0)

    # Step 5: Remove invalid targets
    df, n = remove_invalid_targets(df)
    record("remove_invalid_targets", n)

    # Step 6: Remove negative values in physical columns
    negative_cols = ["weight_lbs", "cubic_feet", "lead_time_hours", "num_carrier_bids"]
    df, n = remove_negative_values(df, negative_cols)
    record("remove_negative_values", n)

    # Step 7: Parse and validate dates
    df, n = remove_invalid_dates(df,
                                  min_date=config.get("min_date", "2023-01-01"),
                                  max_date=config.get("max_date", "2026-12-31"))
    record("remove_invalid_dates", n)

    # Step 8: Drop rows with NaN in any numeric column (no imputation — clean data only)
    before = len(df)
    numeric_cols = config["numeric_columns"] + ["winning_bid_price"]
    df = df.dropna(subset=numeric_cols).copy()
    n_dropped_na = before - len(df)
    record("drop_remaining_nulls", n_dropped_na)
    logger.info(f"drop_remaining_nulls: removed {n_dropped_na} rows with NaN in numeric columns")

    # Step 9: Remove outliers using IQR method
    iqr_cols = config.get("iqr_columns", config["numeric_columns"] + ["winning_bid_price"])
    iqr_multiplier = config.get("iqr_multiplier", 1.5)
    df, n_outliers, outlier_log = remove_outliers_iqr(df, iqr_cols, iqr_multiplier)
    log["outlier_details"] = outlier_log
    record("remove_outliers_iqr", n_outliers)

    log["final_rows"] = len(df)
    log["rows_removed_total"] = log["initial_rows"] - log["final_rows"]

    logger.info(f"Strict cleaning complete: {log['initial_rows']} -> {log['final_rows']} rows ({log['rows_removed_total']} removed)")
    return df, log
