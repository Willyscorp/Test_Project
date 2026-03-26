# Ascent Freight Pricing - Bid Price Prediction

A reusable data pipeline for cleaning, analyzing, and predicting winning bid prices for Ascent Global Logistics freight shipments.

## Objective

Predict `winning_bid_price` from shipment characteristics (weight, equipment type, origin/destination, lead time, etc.).

## Project Structure

```
AscentGlobalLogistics/
├── config.py                  # All pipeline configuration (paths, thresholds, mappings)
├── run_pipeline.py            # Main entry point (CLI)
├── pipeline/
│   ├── clean.py               # Data cleaning functions (9 modular steps)
│   ├── features.py            # Feature matrix preparation
│   └── model.py               # Model training, evaluation, and plotting
├── outputs/
│   ├── ascent_freight_cleaned.csv   # Cleaned dataset
│   ├── model_metrics.json           # Model evaluation results
│   └── plots/                       # Visualization outputs
│       ├── actual_vs_predicted.png
│       ├── feature_importance.png
│       └── residuals.png
├── notebooks/                 # Exploratory analysis notebooks
├── ascent freight pricing dataset.csv   # Raw input data
├── requirements.txt           # Python dependencies
├── create_report.js           # Word document generator (for team reports)
└── Ascent_Freight_Pricing_Data_Assessment_v2.docx  # Latest project documentation
```

## Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

For document generation (optional):
```bash
npm install
```

### 2. Run the Full Pipeline

```bash
python run_pipeline.py
```

This will:
- Load the raw CSV
- Clean the data (handle garbage values, standardize equipment types, impute missing values, cap outliers)
- Train baseline models (Linear Regression + Random Forest)
- Save cleaned data, metrics, and plots to `outputs/`

### 3. Run Individual Steps

```bash
# Only clean the data
python run_pipeline.py --step clean

# Only train models (requires cleaned data from previous step)
python run_pipeline.py --step model

# Run on a different input file (e.g., new data)
python run_pipeline.py --input path/to/new_data.csv
```

## Configuration

All pipeline settings are in `config.py`. Key configurable parameters:

| Parameter | Default | Description |
|-----------|---------|-------------|
| `EQUIPMENT_TYPE_MAPPING` | 32 variants → 7 types | Maps messy equipment names to canonical categories |
| `OUTLIER_THRESHOLDS` | weight: 100K, lead_time: 500, price: 50K | Cap values above these thresholds |
| `IMPUTATION_STRATEGY` | `"median"` | Strategy for filling missing values (`"median"` or `"mean"`) |
| `TRAIN_TEST_SPLIT` | `0.8` | Train/test ratio |
| `RANDOM_SEED` | `42` | Reproducibility seed |

### When New Data Arrives

1. Place the new CSV in the project directory
2. Run: `python run_pipeline.py --input new_data.csv`
3. If there are new equipment type variants, add them to `EQUIPMENT_TYPE_MAPPING` in `config.py`

## Data Cleaning Pipeline

The cleaning pipeline (`pipeline/clean.py`) runs 8 steps in order:

1. **Drop header rows** - Removes duplicate header rows embedded in the data
2. **Drop null rows** - Removes rows where all feature columns are null
3. **Standardize equipment type** - Maps 32 messy variants to 7 canonical types, strips HTML tags
4. **Coerce numeric columns** - Converts garbage values (TBD, #REF!, pending) to NaN
5. **Remove invalid targets** - Drops rows with zero or negative bid prices
6. **Parse dates** - Unifies 3 date formats (YYYY-MM-DD, MM/DD/YYYY, with timestamps)
7. **Cap outliers** - Caps extreme values at configurable thresholds
8. **Impute missing** - Fills NaN values using median (configurable)

**Result:** 9,991 raw rows → 9,934 clean rows (99.4% retention)

## Baseline Model Results

| Model | MAE | RMSE | R² |
|-------|-----|------|-----|
| Linear Regression | $831.49 | $2,457.21 | -0.36 |
| **Random Forest** | **$662.19** | **$1,373.84** | **0.5746** |

- Random Forest explains ~57.5% of variance — a solid baseline
- Linear Regression performs poorly, confirming non-linear relationships in the data
- Feature engineering is planned for the next iteration to improve these numbers

## Documentation

The full project documentation (data quality assessment, pipeline details, actual results, decision log) is in:

**`Ascent_Freight_Pricing_Data_Assessment_v2.docx`**

This is a living document updated after each project phase. Share it with your team for full context.

## Planned Improvements

- [ ] Feature engineering (date features, interactions, lane-level aggregations)
- [ ] Gradient Boosting models (XGBoost / LightGBM)
- [ ] Hyperparameter tuning
- [ ] Cross-validation
- [ ] Time-based train/test split

## Dependencies

- Python 3.10+
- pandas, numpy, scikit-learn, matplotlib, seaborn (see `requirements.txt`)
- Node.js + `docx` package (only for document generation)
