"""
Main entry point for the Ascent Freight Pricing pipeline.

Usage:
    python run_pipeline.py                                    # Full pipeline
    python run_pipeline.py --step clean                       # Only cleaning
    python run_pipeline.py --step model                       # Only modeling (uses existing cleaned data)
    python run_pipeline.py --input new_data.csv               # Run on different input file
"""
import argparse
import logging
import os
import sys

import pandas as pd

import config as cfg
from pipeline.clean import run_cleaning_pipeline
from pipeline.features import build_feature_matrix
from pipeline.model import run_baseline_models, plot_results, save_metrics

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("run_pipeline")


def step_clean(input_path):
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
    }

    df_clean, clean_log = run_cleaning_pipeline(df, cleaning_config)

    os.makedirs(cfg.OUTPUT_DIR, exist_ok=True)
    df_clean.to_csv(cfg.CLEANED_DATA_PATH, index=False)
    logger.info(f"Cleaned data saved to {cfg.CLEANED_DATA_PATH}")

    return df_clean, clean_log


def step_model(df=None):
    """Run the modeling step."""
    if df is None:
        if not os.path.exists(cfg.CLEANED_DATA_PATH):
            logger.error(f"Cleaned data not found at {cfg.CLEANED_DATA_PATH}. Run --step clean first.")
            sys.exit(1)
        logger.info(f"Loading cleaned data from {cfg.CLEANED_DATA_PATH}")
        df = pd.read_csv(cfg.CLEANED_DATA_PATH)

    X, y = build_feature_matrix(df, cfg.TARGET_COLUMN, cfg.DROP_COLUMNS, cfg.CATEGORICAL_COLUMNS)

    results = run_baseline_models(X, y, cfg.TRAIN_TEST_SPLIT, cfg.RANDOM_SEED)

    save_metrics(results["metrics"], cfg.METRICS_PATH)
    plot_results(results, cfg.PLOTS_DIR)

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
    parser.add_argument("--input", default=None,
                        help="Path to input CSV (default: config.RAW_DATA_PATH)")
    args = parser.parse_args()

    input_path = args.input or cfg.RAW_DATA_PATH

    if args.step == "clean":
        step_clean(input_path)
    elif args.step == "model":
        step_model()
    else:
        # Full pipeline
        df_clean, clean_log = step_clean(input_path)
        results = step_model(df_clean)


if __name__ == "__main__":
    main()
