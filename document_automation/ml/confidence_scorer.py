"""
ConfidenceScorer - ML-based document verification confidence scoring.

Uses a Random Forest model trained on crosscheck features to produce
a confidence score (0-100%) indicating how trustworthy a document is.

Features used:
  - Field match ratio (% of fields that match)
  - Text similarity scores for name and address
  - Numeric deviation scores for meter readings and billing
  - Number of anomaly flags
  - Number of missing fields
"""

import numpy as np
import pandas as pd
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import StandardScaler


def _generate_training_data(n_samples: int = 500) -> tuple:
    """
    Generate synthetic training data for the confidence model.

    In production, this would be replaced with historical crosscheck results.
    Here we simulate realistic distributions of valid vs problematic documents.
    """
    rng = np.random.RandomState(42)

    X = []
    y = []

    # --- Valid documents (label=1) ---
    n_valid = int(n_samples * 0.6)
    for _ in range(n_valid):
        match_ratio = rng.uniform(0.85, 1.0)
        name_sim = rng.uniform(0.85, 1.0)
        addr_sim = rng.uniform(0.75, 1.0)
        meter_deviation = rng.uniform(0.0, 0.05)
        billing_deviation = rng.uniform(0.0, 0.1)
        anomaly_count = rng.choice([0, 0, 0, 0, 1], p=[0.4, 0.2, 0.2, 0.1, 0.1])
        missing_fields = rng.choice([0, 0, 1], p=[0.6, 0.3, 0.1])

        X.append([match_ratio, name_sim, addr_sim, meter_deviation,
                  billing_deviation, anomaly_count, missing_fields])
        y.append(1)

    # --- Invalid/suspicious documents (label=0) ---
    n_invalid = n_samples - n_valid
    for _ in range(n_invalid):
        match_ratio = rng.uniform(0.3, 0.85)
        name_sim = rng.uniform(0.2, 0.85)
        addr_sim = rng.uniform(0.2, 0.80)
        meter_deviation = rng.uniform(0.05, 0.5)
        billing_deviation = rng.uniform(0.1, 0.6)
        anomaly_count = rng.choice([1, 2, 3, 4], p=[0.3, 0.3, 0.25, 0.15])
        missing_fields = rng.choice([1, 2, 3, 4], p=[0.3, 0.3, 0.25, 0.15])

        X.append([match_ratio, name_sim, addr_sim, meter_deviation,
                  billing_deviation, anomaly_count, missing_fields])
        y.append(0)

    return np.array(X), np.array(y)


FEATURE_NAMES = [
    "match_ratio",
    "name_similarity",
    "address_similarity",
    "meter_deviation",
    "billing_deviation",
    "anomaly_count",
    "missing_fields",
]


class ConfidenceScorer:
    """
    Score document verification confidence using Random Forest.

    Usage:
        scorer = ConfidenceScorer()
        scorer.train()  # Train on synthetic data (or load historical)
        result = scorer.score(features_dict)
    """

    def __init__(self):
        self.model = RandomForestClassifier(
            n_estimators=100,
            max_depth=8,
            random_state=42,
            class_weight="balanced",
        )
        self.scaler = StandardScaler()
        self._trained = False

    def train(self, X: np.ndarray = None, y: np.ndarray = None) -> dict:
        """
        Train the confidence model.

        Args:
            X: Feature matrix (n_samples, 7). If None, uses synthetic data.
            y: Labels (1=valid, 0=invalid). If None, uses synthetic data.

        Returns:
            dict with training metrics.
        """
        if X is None or y is None:
            X, y = _generate_training_data(500)

        X_scaled = self.scaler.fit_transform(X)
        self.model.fit(X_scaled, y)
        self._trained = True

        # Feature importances
        importances = dict(zip(FEATURE_NAMES, self.model.feature_importances_))

        train_acc = self.model.score(X_scaled, y)

        return {
            "samples": len(y),
            "valid_count": int(np.sum(y == 1)),
            "invalid_count": int(np.sum(y == 0)),
            "train_accuracy": round(train_acc, 4),
            "feature_importances": {k: round(v, 4) for k, v in importances.items()},
        }

    def score(self, features: dict) -> dict:
        """
        Score a single document's verification confidence.

        Args:
            features: dict with keys matching FEATURE_NAMES:
                - match_ratio: float (0-1)
                - name_similarity: float (0-1)
                - address_similarity: float (0-1)
                - meter_deviation: float (0-1, lower is better)
                - billing_deviation: float (0-1, lower is better)
                - anomaly_count: int
                - missing_fields: int

        Returns:
            dict with:
                - confidence_score: float (0-100%)
                - prediction: 'VALID' or 'SUSPICIOUS'
                - risk_level: 'LOW', 'MEDIUM', 'HIGH'
                - feature_contributions: dict
        """
        if not self._trained:
            self.train()

        X = np.array([[
            features.get("match_ratio", 0),
            features.get("name_similarity", 0),
            features.get("address_similarity", 0),
            features.get("meter_deviation", 0.5),
            features.get("billing_deviation", 0.5),
            features.get("anomaly_count", 0),
            features.get("missing_fields", 0),
        ]])

        X_scaled = self.scaler.transform(X)
        probas = self.model.predict_proba(X_scaled)[0]

        # Probability of being valid (class 1)
        confidence = round(float(probas[1]) * 100, 1)

        prediction = "VALID" if confidence >= 60 else "SUSPICIOUS"

        if confidence >= 85:
            risk_level = "LOW"
        elif confidence >= 60:
            risk_level = "MEDIUM"
        else:
            risk_level = "HIGH"

        # Feature contribution (approximate using feature importance * feature value)
        contributions = {}
        importances = self.model.feature_importances_
        for i, name in enumerate(FEATURE_NAMES):
            contributions[name] = round(float(importances[i] * X[0][i]), 4)

        return {
            "confidence_score": confidence,
            "prediction": prediction,
            "risk_level": risk_level,
            "feature_contributions": contributions,
        }

    def score_batch(self, features_list: list[dict]) -> list[dict]:
        """Score multiple documents at once."""
        return [self.score(f) for f in features_list]

    def explain(self) -> dict:
        """Return model details and feature importances for transparency."""
        if not self._trained:
            self.train()

        importances = dict(zip(FEATURE_NAMES, self.model.feature_importances_))
        sorted_imp = sorted(importances.items(), key=lambda x: x[1], reverse=True)

        return {
            "model_type": "RandomForestClassifier",
            "n_estimators": self.model.n_estimators,
            "max_depth": self.model.max_depth,
            "feature_importances_ranked": [
                {"feature": k, "importance": round(v, 4)} for k, v in sorted_imp
            ],
        }
