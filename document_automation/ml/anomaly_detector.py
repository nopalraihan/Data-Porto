"""
AnomalyDetector - Detect suspicious values in PLN document fields.

Uses Isolation Forest and statistical rules to flag:
  - Abnormal kWh consumption (too high or too low for the tariff)
  - Billing amounts that don't match consumption patterns
  - Meter reading inconsistencies (akhir < awal, impossible jumps)
  - Unusual patterns suggesting data entry errors or fraud
"""

import numpy as np
import pandas as pd
from sklearn.ensemble import IsolationForest
from sklearn.preprocessing import StandardScaler


# Expected consumption ranges by tariff type (kWh per month)
TARIFF_CONSUMPTION_RANGES = {
    "R1/450": (20, 150),
    "R1/900": (50, 300),
    "R1M/900": (50, 300),
    "R1/1300": (80, 500),
    "R1/2200": (100, 800),
    "R2/3500": (150, 1500),
    "R2/5500": (200, 2500),
    "R3/6600": (300, 5000),
    "B1/1300": (100, 1000),
    "B1/2200": (150, 1500),
    "B2/6600": (300, 5000),
}

# Approximate rate per kWh by tariff (Rupiah)
TARIFF_RATE_APPROX = {
    "R1/450": 415,
    "R1/900": 605,
    "R1M/900": 1352,
    "R1/1300": 1444,
    "R1/2200": 1444,
    "R2/3500": 1699,
    "R2/5500": 1699,
    "R3/6600": 1699,
    "B1/1300": 1444,
    "B1/2200": 1444,
    "B2/6600": 1444,
}


class AnomalyDetector:
    """
    Detect anomalies in PLN document data.

    Usage:
        detector = AnomalyDetector()
        flags = detector.check_single(fields_dict)
        flags = detector.check_batch(dataframe)
    """

    def __init__(self, contamination: float = 0.1):
        """
        Args:
            contamination: Expected proportion of anomalies (for Isolation Forest).
        """
        self.contamination = contamination

    def check_single(self, fields: dict) -> list[dict]:
        """
        Check a single record for anomalies.

        Args:
            fields: dict with keys like 'tarif_daya', 'stand_meter_awal',
                    'stand_meter_akhir', 'pemakaian_kwh', 'biaya_listrik'

        Returns:
            List of anomaly flags, each a dict with:
                - field: field name
                - severity: 'WARNING' or 'CRITICAL'
                - message: description
                - expected: expected range/value
                - actual: actual value
        """
        flags = []

        stand_awal = self._to_float(fields.get("stand_meter_awal"))
        stand_akhir = self._to_float(fields.get("stand_meter_akhir"))
        pemakaian = self._to_float(fields.get("pemakaian_kwh"))
        biaya = self._to_float(fields.get("biaya_listrik"))
        tarif = str(fields.get("tarif_daya", "")).strip().upper()

        # --- Rule 1: Meter reading consistency ---
        if stand_awal is not None and stand_akhir is not None:
            if stand_akhir < stand_awal:
                flags.append({
                    "field": "stand_meter_akhir",
                    "severity": "CRITICAL",
                    "message": "Stand meter akhir is LESS than stand meter awal (meter went backwards)",
                    "expected": f"> {stand_awal}",
                    "actual": str(stand_akhir),
                })

            # Check if pemakaian matches the meter difference
            expected_usage = stand_akhir - stand_awal
            if pemakaian is not None and abs(expected_usage - pemakaian) > 1:
                flags.append({
                    "field": "pemakaian_kwh",
                    "severity": "CRITICAL",
                    "message": f"Pemakaian ({pemakaian}) doesn't match meter difference ({expected_usage})",
                    "expected": str(expected_usage),
                    "actual": str(pemakaian),
                })

        # --- Rule 2: Consumption vs tariff range ---
        if pemakaian is not None and tarif:
            tarif_key = self._normalize_tarif(tarif)
            if tarif_key in TARIFF_CONSUMPTION_RANGES:
                low, high = TARIFF_CONSUMPTION_RANGES[tarif_key]
                if pemakaian < low * 0.5:
                    flags.append({
                        "field": "pemakaian_kwh",
                        "severity": "WARNING",
                        "message": f"Unusually LOW consumption for tariff {tarif_key}",
                        "expected": f"{low} - {high} kWh",
                        "actual": f"{pemakaian} kWh",
                    })
                elif pemakaian > high * 1.5:
                    flags.append({
                        "field": "pemakaian_kwh",
                        "severity": "WARNING",
                        "message": f"Unusually HIGH consumption for tariff {tarif_key}",
                        "expected": f"{low} - {high} kWh",
                        "actual": f"{pemakaian} kWh",
                    })

        # --- Rule 3: Billing vs consumption consistency ---
        if pemakaian is not None and biaya is not None and pemakaian > 0:
            rate = biaya / pemakaian
            tarif_key = self._normalize_tarif(tarif)
            expected_rate = TARIFF_RATE_APPROX.get(tarif_key)

            if expected_rate:
                deviation = abs(rate - expected_rate) / expected_rate
                if deviation > 0.3:  # More than 30% deviation
                    flags.append({
                        "field": "biaya_listrik",
                        "severity": "WARNING",
                        "message": f"Billing rate Rp {rate:,.0f}/kWh deviates {deviation*100:.0f}% from expected Rp {expected_rate}/kWh",
                        "expected": f"~Rp {expected_rate}/kWh (total ~Rp {expected_rate * pemakaian:,.0f})",
                        "actual": f"Rp {rate:,.0f}/kWh (total Rp {biaya:,.0f})",
                    })

        # --- Rule 4: Zero/negative checks ---
        if pemakaian is not None and pemakaian <= 0:
            flags.append({
                "field": "pemakaian_kwh",
                "severity": "WARNING",
                "message": "Zero or negative consumption",
                "expected": "> 0",
                "actual": str(pemakaian),
            })

        if biaya is not None and biaya <= 0:
            flags.append({
                "field": "biaya_listrik",
                "severity": "WARNING",
                "message": "Zero or negative billing amount",
                "expected": "> 0",
                "actual": str(biaya),
            })

        return flags

    def check_batch(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Run Isolation Forest anomaly detection on a DataFrame of records.

        Expects columns: stand_meter_awal, stand_meter_akhir, pemakaian_kwh, biaya_listrik

        Returns:
            DataFrame with added columns:
                - anomaly_score: float (-1 = anomaly, 1 = normal)
                - is_anomaly: bool
                - anomaly_flags: list of per-row flags
        """
        numeric_cols = []
        col_map = {
            "Stand Meter Awal": "stand_meter_awal",
            "Stand Meter Akhir": "stand_meter_akhir",
            "Pemakaian (kWh)": "pemakaian_kwh",
            "Biaya Listrik (Rp)": "biaya_listrik",
        }

        result_df = df.copy()

        # Prepare numeric features for Isolation Forest
        feature_cols = []
        for excel_col, field_key in col_map.items():
            if excel_col in df.columns:
                result_df[f"_num_{field_key}"] = pd.to_numeric(
                    df[excel_col].astype(str).str.replace(r"[,.\s]", "", regex=True),
                    errors="coerce"
                )
                feature_cols.append(f"_num_{field_key}")

        if len(feature_cols) >= 2 and len(df) >= 5:
            features = result_df[feature_cols].fillna(0).values
            scaler = StandardScaler()
            features_scaled = scaler.fit_transform(features)

            iso_forest = IsolationForest(
                contamination=self.contamination,
                random_state=42,
                n_estimators=100,
            )
            result_df["anomaly_score"] = iso_forest.fit_predict(features_scaled)
            result_df["is_anomaly"] = result_df["anomaly_score"] == -1
        else:
            result_df["anomaly_score"] = 1
            result_df["is_anomaly"] = False

        # Run rule-based checks per row
        all_flags = []
        for idx in range(len(df)):
            row = df.iloc[idx]
            fields = {}
            for excel_col, field_key in col_map.items():
                if excel_col in row.index:
                    fields[field_key] = row[excel_col]
            # Add tarif
            tarif_col = "Tarif/Daya" if "Tarif/Daya" in row.index else None
            if tarif_col:
                fields["tarif_daya"] = row[tarif_col]

            flags = self.check_single(fields)
            all_flags.append(flags)

        result_df["anomaly_flags"] = all_flags

        # Clean up temp columns
        for col in feature_cols:
            result_df.drop(columns=col, inplace=True, errors="ignore")

        return result_df

    @staticmethod
    def _to_float(value) -> float | None:
        """Convert a value to float, handling common formats."""
        if value is None:
            return None
        s = str(value).strip()
        s = s.replace(",", "").replace(".", "").replace(" ", "")
        s = s.replace("Rp", "").replace("rp", "").replace("kWh", "").replace("kwh", "")
        try:
            return float(s)
        except ValueError:
            return None

    @staticmethod
    def _normalize_tarif(tarif: str) -> str:
        """Normalize tariff string to match lookup keys."""
        t = tarif.strip().upper()
        t = t.replace(" ", "").replace("-", "/")
        # Ensure format like R1/1300
        t = t.replace("VA", "").replace("W", "").strip().rstrip("/")
        return t
