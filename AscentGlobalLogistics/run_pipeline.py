"""
Main entry point for the Ascent Freight Pricing pipeline.

Usage:
    python run_pipeline.py                                    # Full pipeline (baseline mode)
    python run_pipeline.py --mode strict                      # Strict mode (remove outliers + garbage)
    python run_pipeline.py --step clean                       # Only cleaning
    python run_pipeline.py --step model                       # Only modeling (uses existing cleaned data)
    python run_pipeline.py --input new_data.csv               # Run on different input file
"""
import argparse
import json
import logging
import os
import sys

import pandas as pd

import config as cfg
from pipeline.clean import run_cleaning_pipeline, run_strict_cleaning_pipeline
from pipeline.features import build_feature_matrix
from pipeline.model import run_baseline_models, plot_results, save_metrics

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("run_pipeline")


def step_clean(input_path, mode="baseline"):
    """Run the cleaning step and save cleaned data."""
    logger.info(f"Loading raw data from {input_path}")
    df = pd.read_csv(input_path)
    logger.info(f"Raw data: {df.shape[0]} rows, {df.shape[1]} columns")

    cleaning_config = {
        "equipment_mapping": cfg.EQUIPMENT_TYPE_MAPPING,
        "equipment_drop_values": cfg.EQUIPMENT_TYPE_DROP_VALUES,
        "numeric_columns": cfg.NUMERIC_COLUMNS,
        "outlier_thresholds": cfg.OUTLIER_THRESHOLDS,
        "imputation_strategy": cfg.IMPUTATION_STRATEGY,
        "min_date": "2023-01-01",
        "max_date": "2026-12-31",
        "iqr_columns": cfg.NUMERIC_COLUMNS + [cfg.TARGET_COLUMN],
        "iqr_multiplier": 1.5,
    }

    if mode == "strict":
        output_dir = os.path.join(cfg.OUTPUT_DIR, "strict")
        cleaned_path = os.path.join(output_dir, "ascent_freight_cleaned_strict.csv")
        df_clean, clean_log = run_strict_cleaning_pipeline(df, cleaning_config)
    else:
        output_dir = cfg.OUTPUT_DIR
        cleaned_path = cfg.CLEANED_DATA_PATH
        df_clean, clean_log = run_cleaning_pipeline(df, cleaning_config)

    os.makedirs(output_dir, exist_ok=True)
    df_clean.to_csv(cleaned_path, index=False)
    logger.info(f"Cleaned data saved to {cleaned_path}")

    # Save cleaning log
    log_path = os.path.join(output_dir, "cleaning_log.json")
    with open(log_path, "w") as f:
        json.dump(clean_log, f, indent=2, default=str)

    return df_clean, clean_log, output_dir


def step_model(df=None, output_dir=None):
    """Run the modeling step."""
    output_dir = output_dir or cfg.OUTPUT_DIR
    cleaned_path = os.path.join(output_dir, "ascent_freight_cleaned_strict.csv") \
        if "strict" in output_dir else cfg.CLEANED_DATA_PATH

    if df is None:
        if not os.path.exists(cleaned_path):
            logger.error(f"Cleaned data not found at {cleaned_path}. Run --step clean first.")
            sys.exit(1)
        logger.info(f"Loading cleaned data from {cleaned_path}")
        df = pd.read_csv(cleaned_path)

    X, y = build_feature_matrix(df, cfg.TARGET_COLUMN, cfg.DROP_COLUMNS, cfg.CATEGORICAL_COLUMNS)

    results = run_baseline_models(X, y, cfg.TRAIN_TEST_SPLIT, cfg.RANDOM_SEED)

    metrics_path = os.path.join(output_dir, "model_metrics.json")
    plots_dir = os.path.join(output_dir, "plots")
    save_metrics(results["metrics"], metrics_path)
    plot_results(results, plots_dir)

    print("\n" + "=" * 60)
    print("BASELINE MODEL RESULTS")
    print("=" * 60)
    for m in results["metrics"]:
        print(f"\n  {m['model']}:")
        print(f"    MAE  = ${m['MAE']:,.2f}")
        print(f"    RMSE = ${m['RMSE']:,.2f}")
        print(f"    R²   = {m['R2']:.4f}")
    print("=" * 60)

    return results


def main():
    parser = argparse.ArgumentParser(description="Ascent Freight Pricing Pipeline")
    parser.add_argument("--step", choices=["clean", "model"], default=None,
                        help="Run a specific step (default: full pipeline)")
    parser.add_argument("--mode", choices=["baseline", "strict"], default="baseline",
                        help="Cleaning mode: baseline (cap outliers) or strict (remove outliers)")
    parser.add_argument("--input", default=None,
                        help="Path to input CSV (default: config.RAW_DATA_PATH)")
    args = parser.parse_args()

    input_path = args.input or cfg.RAW_DATA_PATH

    if args.step == "clean":
        step_clean(input_path, args.mode)
    elif args.step == "model":
        output_dir = os.path.join(cfg.OUTPUT_DIR, "strict") if args.mode == "strict" else cfg.OUTPUT_DIR
        step_model(output_dir=output_dir)
    else:
        # Full pipeline
        df_clean, clean_log, output_dir = step_clean(input_path, args.mode)
        results = step_model(df_clean, output_dir)


if __name__ == "__main__":
    main()
