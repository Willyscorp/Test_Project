"""
Feature preparation for Ascent Freight Pricing.
NOTE: Feature engineering is deferred to a future iteration.
This module provides minimal preparation for the baseline model.
"""
import logging

import pandas as pd

logger = logging.getLogger(__name__)


def build_feature_matrix(df, target_col, drop_cols, categorical_cols):
    """
    Minimal feature preparation for baseline model:
    - Drop non-feature columns (shipment_id, shipment_date)
    - One-hot encode categoricals
    - Return X, y
    """
    df = df.copy()

    # Separate target
    y = df[target_col].copy()

    # Drop target and non-feature columns
    cols_to_drop = [c for c in drop_cols + [target_col] if c in df.columns]
    df = df.drop(columns=cols_to_drop)

    # One-hot encode categoricals
    df = pd.get_dummies(df, columns=categorical_cols, drop_first=True, dtype=int)

    logger.info(f"build_feature_matrix: {df.shape[1]} features, {len(y)} samples")
    return df, y
