"""
Pipeline configuration for Ascent Freight Pricing.
All thresholds, mappings, and paths are centralized here.
"""
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# --- File paths ---
RAW_DATA_PATH = os.path.join(BASE_DIR, "ascent freight pricing dataset.csv")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
CLEANED_DATA_PATH = os.path.join(OUTPUT_DIR, "ascent_freight_cleaned.csv")
METRICS_PATH = os.path.join(OUTPUT_DIR, "model_metrics.json")
PLOTS_DIR = os.path.join(OUTPUT_DIR, "plots")

# --- Numeric columns ---
NUMERIC_COLUMNS = ["weight_lbs", "cubic_feet", "lead_time_hours", "num_carrier_bids", "lane_volume"]
TARGET_COLUMN = "winning_bid_price"

# --- Categorical columns ---
CATEGORICAL_COLUMNS = ["origin_region", "destination_region", "equipment_type"]

# --- Columns to drop before modeling ---
DROP_COLUMNS = ["shipment_id", "shipment_date"]

# --- Equipment type mapping (32 variants -> 7 canonical types) ---
EQUIPMENT_TYPE_MAPPING = {
    # Cargo Van
    "Cargo Van": "Cargo Van",
    "cargo van": "Cargo Van",
    "CARGO VAN": "Cargo Van",
    "CV": "Cargo Van",
    "Cargo  Van": "Cargo Van",
    "<b>Cargo Van</b>": "Cargo Van",

    # Sprinter Van
    "Sprinter Van": "Sprinter Van",
    "sprinter van": "Sprinter Van",
    "SPRINTER VAN": "Sprinter Van",
    "Sprinter": "Sprinter Van",
    "Sprinter Van\xa0": "Sprinter Van",

    # Small Straight
    "Small Straight": "Small Straight",
    "small straight": "Small Straight",
    "Sm Straight": "Small Straight",
    "Small Str": "Small Straight",

    # Large Straight
    "Large Straight": "Large Straight",
    "LARGE STRAIGHT": "Large Straight",
    "Large Str": "Large Straight",
    "Lg Straight": "Large Straight",

    # Tractor Trailer
    "Tractor Trailer": "Tractor Trailer",
    "tractor trailer": "Tractor Trailer",
    "Tractor": "Tractor Trailer",
    "TT": "Tractor Trailer",

    # Flatbed
    "Flatbed": "Flatbed",
    "flatbed": "Flatbed",
    "FLATBED": "Flatbed",
    "Flat Bed": "Flatbed",

    # Reefer
    "Reefer": "Reefer",
    "reefer": "Reefer",
    "REEFER": "Reefer",
    "Refrigerated": "Reefer",
}

# Values to drop (not mappable to a canonical type)
EQUIPMENT_TYPE_DROP_VALUES = ["Test Equipment", "equipment_type"]

# --- Outlier thresholds ---
OUTLIER_THRESHOLDS = {
    "weight_lbs": 100_000,
    "cubic_feet": 5_000,
    "lead_time_hours": 500,
    "winning_bid_price": 50_000,
}

# --- Imputation ---
IMPUTATION_STRATEGY = "median"  # "median" or "mean"

# --- Model settings ---
TRAIN_TEST_SPLIT = 0.8
RANDOM_SEED = 42
