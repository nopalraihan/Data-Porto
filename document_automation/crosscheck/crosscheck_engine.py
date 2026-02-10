"""
CrosscheckEngine - Compares extracted PDF data against Excel template records.

ML-enhanced with:
  - TF-IDF fuzzy matching for names/addresses
  - Anomaly detection for meter readings and billing
  - Random Forest confidence scoring
"""

import re
from datetime import datetime

import pandas as pd

from document_automation.ml.text_similarity import TextSimilarity
from document_automation.ml.anomaly_detector import AnomalyDetector
from document_automation.ml.confidence_scorer import ConfidenceScorer


# Maps Excel column names to PDF field keys
COLUMN_TO_FIELD = {
    "ID Pelanggan": "id_pelanggan",
    "Nama Pelanggan": "nama_pelanggan",
    "Alamat": "alamat",
    "Tarif/Daya": "tarif_daya",
    "Nomor Meter": "nomor_meter",
    "Nomor kWh": "nomor_kwh",
    "Periode": "periode",
    "Stand Meter Awal": "stand_meter_awal",
    "Stand Meter Akhir": "stand_meter_akhir",
    "Pemakaian (kWh)": "pemakaian_kwh",
    "Biaya Listrik (Rp)": "biaya_listrik",
}

# Fields to crosscheck (excludes No and Status which are metadata columns)
CROSSCHECK_FIELDS = list(COLUMN_TO_FIELD.keys())


def _normalize(value) -> str:
    """Normalize a value for comparison: strip, uppercase, remove extra spaces/punctuation."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).strip().upper()
    # Remove dots, commas, dashes in numbers for numeric comparison
    s = re.sub(r"\s+", " ", s)
    return s


def _normalize_numeric(value) -> str:
    """Normalize numeric values by removing separators."""
    s = _normalize(value)
    # Remove Rp, dots, commas for numeric fields
    s = re.sub(r"[Rr][Pp]\.?\s*", "", s)
    s = re.sub(r"[.,\s]", "", s)
    return s


def _is_numeric_field(field_name: str) -> bool:
    """Check if a field should be compared numerically."""
    numeric_fields = {
        "Stand Meter Awal", "Stand Meter Akhir",
        "Pemakaian (kWh)", "Biaya Listrik (Rp)",
    }
    return field_name in numeric_fields


def _fuzzy_contains(pdf_val: str, excel_val: str) -> bool:
    """Check if one string contains the other (for address/name partial matching)."""
    a = _normalize(pdf_val)
    b = _normalize(excel_val)
    if not a or not b:
        return False
    return a in b or b in a


class CrosscheckEngine:
    """
    ML-enhanced crosscheck engine.

    Compare PDF-extracted fields against Excel template rows using:
      - TF-IDF cosine similarity for fuzzy text matching
      - Isolation Forest for anomaly detection
      - Random Forest for overall confidence scoring

    Usage:
        engine = CrosscheckEngine(pdf_fields, excel_df)
        results = engine.run()
    """

    # Text fields that benefit from ML similarity matching
    TEXT_FIELDS = {"Nama Pelanggan", "Alamat"}

    def __init__(self, pdf_fields: dict, excel_df: pd.DataFrame, use_ml: bool = True):
        """
        Args:
            pdf_fields: dict of field_name -> value (from PLNExtractor.extract_flat())
            excel_df: DataFrame from the 'Data Pelanggan' sheet of the template
            use_ml: Enable ML-enhanced matching (default True)
        """
        self.pdf_fields = pdf_fields
        self.excel_df = excel_df
        self.use_ml = use_ml

        # Initialize ML components
        if self.use_ml:
            self.text_sim = TextSimilarity()
            self.anomaly_detector = AnomalyDetector()
            self.confidence_scorer = ConfidenceScorer()
            self.confidence_scorer.train()

    def run(self) -> dict:
        """
        Run crosscheck and return results.

        Returns dict with:
            - 'matched_row': index of best-matching Excel row (or None)
            - 'results': list of per-field comparison dicts
            - 'summary': overall match statistics
        """
        matched_row_idx = self._find_matching_row()

        if matched_row_idx is None:
            return self._no_match_result()

        excel_row = self.excel_df.iloc[matched_row_idx]
        results = []

        for excel_col, pdf_key in COLUMN_TO_FIELD.items():
            pdf_value = self.pdf_fields.get(pdf_key)
            excel_value = excel_row.get(excel_col)

            comparison = self._compare_field(excel_col, pdf_key, pdf_value, excel_value)
            results.append(comparison)

        # Compute summary
        total = len(results)
        matches = sum(1 for r in results if r["match_status"] == "MATCH")
        mismatches = sum(1 for r in results if r["match_status"] == "MISMATCH")
        missing = sum(1 for r in results if r["match_status"] == "MISSING")

        summary = {
            "matched_row_index": matched_row_idx,
            "matched_row_number": matched_row_idx + 1,
            "total_fields_checked": total,
            "total_match": matches,
            "total_mismatch": mismatches,
            "total_missing": missing,
            "match_percentage": round((matches / total) * 100, 1) if total > 0 else 0,
            "checked_at": datetime.now().isoformat(),
        }

        output = {
            "matched_row": matched_row_idx,
            "results": results,
            "summary": summary,
        }

        # ── ML Enhancements ──
        if self.use_ml:
            # 1. Compute text similarity scores
            name_sim = self._ml_text_similarity("nama_pelanggan", "Nama Pelanggan", excel_row)
            addr_sim = self._ml_text_similarity("alamat", "Alamat", excel_row)
            output["ml_similarity"] = {
                "nama_pelanggan": name_sim,
                "alamat": addr_sim,
            }

            # Enhance match results with similarity scores
            for r in results:
                if r["field_name"] in self.TEXT_FIELDS:
                    key = COLUMN_TO_FIELD[r["field_name"]]
                    sim = name_sim if key == "nama_pelanggan" else addr_sim
                    r["similarity_score"] = sim["score"]
                    r["similarity_class"] = sim["classification"]

            # 2. Anomaly detection on the PDF data
            anomaly_flags = self.anomaly_detector.check_single(self.pdf_fields)
            output["ml_anomalies"] = anomaly_flags

            # 3. Confidence scoring
            meter_dev = self._compute_meter_deviation(excel_row)
            billing_dev = self._compute_billing_deviation(excel_row)

            confidence_features = {
                "match_ratio": matches / total if total > 0 else 0,
                "name_similarity": name_sim["score"],
                "address_similarity": addr_sim["score"],
                "meter_deviation": meter_dev,
                "billing_deviation": billing_dev,
                "anomaly_count": len(anomaly_flags),
                "missing_fields": missing,
            }
            confidence_result = self.confidence_scorer.score(confidence_features)
            output["ml_confidence"] = confidence_result

        return output

    def run_all_rows(self) -> list[dict]:
        """
        Compare PDF fields against ALL rows in the Excel template.
        Returns a list of results, one per Excel row.
        """
        all_results = []
        for idx in range(len(self.excel_df)):
            excel_row = self.excel_df.iloc[idx]
            row_results = []

            for excel_col, pdf_key in COLUMN_TO_FIELD.items():
                pdf_value = self.pdf_fields.get(pdf_key)
                excel_value = excel_row.get(excel_col)
                comparison = self._compare_field(excel_col, pdf_key, pdf_value, excel_value)
                row_results.append(comparison)

            total = len(row_results)
            matches = sum(1 for r in row_results if r["match_status"] == "MATCH")

            all_results.append({
                "row_index": idx,
                "row_number": idx + 1,
                "results": row_results,
                "match_count": matches,
                "match_percentage": round((matches / total) * 100, 1) if total > 0 else 0,
            })

        return all_results

    def _find_matching_row(self) -> int | None:
        """
        Find the Excel row that best matches the PDF.

        Uses exact matching for IDs and meter numbers, plus ML-based
        TF-IDF similarity for names and addresses when exact match fails.
        """
        pdf_id = _normalize_numeric(self.pdf_fields.get("id_pelanggan", ""))
        pdf_meter = _normalize(self.pdf_fields.get("nomor_meter", ""))
        pdf_name = self.pdf_fields.get("nama_pelanggan", "")
        pdf_addr = self.pdf_fields.get("alamat", "")

        best_idx = None
        best_score = 0.0

        for idx in range(len(self.excel_df)):
            row = self.excel_df.iloc[idx]
            score = 0.0

            # Match by ID Pelanggan (exact, highest weight)
            excel_id = _normalize_numeric(row.get("ID Pelanggan", ""))
            if pdf_id and excel_id and pdf_id == excel_id:
                score += 10

            # Match by Nomor Meter (exact)
            excel_meter = _normalize(row.get("Nomor Meter", ""))
            if pdf_meter and excel_meter and pdf_meter == excel_meter:
                score += 5

            # ML-enhanced name matching
            excel_name = str(row.get("Nama Pelanggan", ""))
            if pdf_name and excel_name:
                if self.use_ml:
                    sim = self.text_sim.score(pdf_name, excel_name)
                    score += sim * 4  # Up to 4 points based on similarity
                elif _fuzzy_contains(pdf_name, excel_name):
                    score += 3

            # ML-enhanced address matching
            excel_addr = str(row.get("Alamat", ""))
            if pdf_addr and excel_addr and self.use_ml:
                sim = self.text_sim.score(pdf_addr, excel_addr)
                score += sim * 2  # Up to 2 points

            if score > best_score:
                best_score = score
                best_idx = idx

        return best_idx if best_score >= 3 else None

    def _compare_field(self, excel_col: str, pdf_key: str, pdf_value, excel_value) -> dict:
        """Compare a single field between PDF and Excel."""
        result = {
            "field_name": excel_col,
            "pdf_key": pdf_key,
            "pdf_value": str(pdf_value) if pdf_value else None,
            "excel_value": str(excel_value) if excel_value and not (isinstance(excel_value, float) and pd.isna(excel_value)) else None,
        }

        # Both missing
        if not result["pdf_value"] and not result["excel_value"]:
            result["match_status"] = "MISSING"
            result["notes"] = "Both PDF and Excel values are empty"
            return result

        # One side missing
        if not result["pdf_value"]:
            result["match_status"] = "MISSING"
            result["notes"] = "Not found in PDF document"
            return result
        if not result["excel_value"]:
            result["match_status"] = "MISSING"
            result["notes"] = "Not found in Excel template"
            return result

        # Numeric comparison
        if _is_numeric_field(excel_col):
            pdf_norm = _normalize_numeric(pdf_value)
            excel_norm = _normalize_numeric(excel_value)
            if pdf_norm == excel_norm:
                result["match_status"] = "MATCH"
                result["notes"] = "Exact numeric match"
            else:
                result["match_status"] = "MISMATCH"
                result["notes"] = f"PDF='{pdf_norm}' vs Excel='{excel_norm}'"
            return result

        # String comparison (ML-enhanced for text fields)
        pdf_norm = _normalize(pdf_value)
        excel_norm = _normalize(excel_value)

        if pdf_norm == excel_norm:
            result["match_status"] = "MATCH"
            result["notes"] = "Exact match"
        elif _fuzzy_contains(pdf_norm, excel_norm):
            result["match_status"] = "MATCH"
            result["notes"] = "Partial/contains match"
        elif self.use_ml and excel_col in self.TEXT_FIELDS:
            sim_score = self.text_sim.score(str(pdf_value), str(excel_value))
            sim_class = self.text_sim.classify_match(sim_score)
            if sim_score >= 0.75:
                result["match_status"] = "MATCH"
                result["notes"] = f"ML similarity match ({sim_class}, score={sim_score:.2f})"
            else:
                result["match_status"] = "MISMATCH"
                result["notes"] = f"ML similarity={sim_score:.2f} ({sim_class}). PDF='{pdf_value}' vs Excel='{excel_value}'"
        else:
            result["match_status"] = "MISMATCH"
            result["notes"] = f"PDF='{pdf_value}' vs Excel='{excel_value}'"

        return result

    def _ml_text_similarity(self, pdf_key: str, excel_col: str, excel_row) -> dict:
        """Compute ML text similarity between a PDF field and Excel column."""
        pdf_val = str(self.pdf_fields.get(pdf_key, ""))
        excel_val = str(excel_row.get(excel_col, ""))

        if not pdf_val or not excel_val:
            return {"score": 0.0, "classification": "NO_MATCH"}

        score = self.text_sim.score(pdf_val, excel_val)
        classification = self.text_sim.classify_match(score)
        return {"score": score, "classification": classification}

    def _compute_meter_deviation(self, excel_row) -> float:
        """Compute deviation between PDF and Excel meter readings (0=perfect, 1=max deviation)."""
        pdf_akhir = _normalize_numeric(self.pdf_fields.get("stand_meter_akhir", "0"))
        excel_akhir = _normalize_numeric(excel_row.get("Stand Meter Akhir", "0"))
        try:
            pdf_v = float(pdf_akhir) if pdf_akhir else 0
            excel_v = float(excel_akhir) if excel_akhir else 0
            if excel_v == 0:
                return 0.0
            return min(abs(pdf_v - excel_v) / excel_v, 1.0)
        except (ValueError, ZeroDivisionError):
            return 0.5

    def _compute_billing_deviation(self, excel_row) -> float:
        """Compute deviation between PDF and Excel billing amounts."""
        pdf_biaya = _normalize_numeric(self.pdf_fields.get("biaya_listrik", "0"))
        excel_biaya = _normalize_numeric(excel_row.get("Biaya Listrik (Rp)", "0"))
        try:
            pdf_v = float(pdf_biaya) if pdf_biaya else 0
            excel_v = float(excel_biaya) if excel_biaya else 0
            if excel_v == 0:
                return 0.0
            return min(abs(pdf_v - excel_v) / excel_v, 1.0)
        except (ValueError, ZeroDivisionError):
            return 0.5

    def _no_match_result(self) -> dict:
        """Return result when no matching row is found in Excel."""
        results = []
        for excel_col, pdf_key in COLUMN_TO_FIELD.items():
            pdf_value = self.pdf_fields.get(pdf_key)
            results.append({
                "field_name": excel_col,
                "pdf_key": pdf_key,
                "pdf_value": str(pdf_value) if pdf_value else None,
                "excel_value": None,
                "match_status": "MISSING",
                "notes": "No matching row found in Excel template",
            })

        return {
            "matched_row": None,
            "results": results,
            "summary": {
                "matched_row_index": None,
                "matched_row_number": None,
                "total_fields_checked": len(results),
                "total_match": 0,
                "total_mismatch": 0,
                "total_missing": len(results),
                "match_percentage": 0,
                "checked_at": datetime.now().isoformat(),
            },
        }
