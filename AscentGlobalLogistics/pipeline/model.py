"""
Model training and evaluation for Ascent Freight Pricing.
"""
import json
import logging
import os

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import RandomForestRegressor
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_absolute_error, root_mean_squared_error, r2_score

logger = logging.getLogger(__name__)


def train_evaluate_model(X_train, X_test, y_train, y_test, model, name):
    """Train a model and return metrics dict."""
    model.fit(X_train, y_train)
    y_pred = model.predict(X_test)

    metrics = {
        "model": name,
        "MAE": round(mean_absolute_error(y_test, y_pred), 2),
        "RMSE": round(root_mean_squared_error(y_test, y_pred), 2),
        "R2": round(r2_score(y_test, y_pred), 4),
    }
    logger.info(f"{name}: MAE={metrics['MAE']}, RMSE={metrics['RMSE']}, R²={metrics['R2']}")
    return metrics, y_pred, model


def run_baseline_models(X, y, test_size, random_seed):
    """Split data and run baseline models. Returns results dict."""
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=1 - test_size, random_state=random_seed
    )
    logger.info(f"Train/test split: {len(X_train)} train, {len(X_test)} test")

    models = [
        ("Linear Regression", LinearRegression()),
        ("Random Forest", RandomForestRegressor(n_estimators=100, random_state=random_seed, n_jobs=-1)),
    ]

    results = []
    predictions = {}
    trained_models = {}

    for name, model in models:
        metrics, y_pred, fitted = train_evaluate_model(X_train, X_test, y_train, y_test, model, name)
        results.append(metrics)
        predictions[name] = y_pred
        trained_models[name] = fitted

    return {
        "metrics": results,
        "predictions": predictions,
        "trained_models": trained_models,
        "X_test": X_test,
        "y_test": y_test,
        "feature_names": list(X.columns),
    }


def plot_results(results, output_dir):
    """Generate and save evaluation plots."""
    os.makedirs(output_dir, exist_ok=True)
    y_test = results["y_test"]

    # Actual vs Predicted for each model
    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    for ax, (name, y_pred) in zip(axes, results["predictions"].items()):
        ax.scatter(y_test, y_pred, alpha=0.3, s=10)
        ax.plot([y_test.min(), y_test.max()], [y_test.min(), y_test.max()], "r--", lw=2)
        ax.set_xlabel("Actual Price ($)")
        ax.set_ylabel("Predicted Price ($)")
        ax.set_title(f"{name}: Actual vs Predicted")
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "actual_vs_predicted.png"), dpi=150)
    plt.close()

    # Residuals
    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    for ax, (name, y_pred) in zip(axes, results["predictions"].items()):
        residuals = y_test.values - y_pred
        ax.scatter(y_pred, residuals, alpha=0.3, s=10)
        ax.axhline(y=0, color="r", linestyle="--", lw=2)
        ax.set_xlabel("Predicted Price ($)")
        ax.set_ylabel("Residual ($)")
        ax.set_title(f"{name}: Residuals")
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "residuals.png"), dpi=150)
    plt.close()

    # Feature importance (Random Forest)
    if "Random Forest" in results["trained_models"]:
        rf = results["trained_models"]["Random Forest"]
        importances = rf.feature_importances_
        feature_names = results["feature_names"]
        indices = np.argsort(importances)[-20:]  # Top 20

        fig, ax = plt.subplots(figsize=(10, 8))
        ax.barh(range(len(indices)), importances[indices])
        ax.set_yticks(range(len(indices)))
        ax.set_yticklabels([feature_names[i] for i in indices])
        ax.set_xlabel("Feature Importance")
        ax.set_title("Random Forest: Top 20 Feature Importances")
        plt.tight_layout()
        plt.savefig(os.path.join(output_dir, "feature_importance.png"), dpi=150)
        plt.close()

    logger.info(f"Plots saved to {output_dir}")


def save_metrics(metrics, path):
    """Save metrics list to JSON."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w") as f:
        json.dump(metrics, f, indent=2)
    logger.info(f"Metrics saved to {path}")
