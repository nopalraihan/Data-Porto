"""
CrosscheckEngine - Compares extracted PDF data against Excel template records.
"""

import re
from datetime import datetime

import pandas as pd


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
    Compare PDF-extracted fields against Excel template rows.

    Usage:
        engine = CrosscheckEngine(pdf_fields, excel_df)
        results = engine.run()
    """

    def __init__(self, pdf_fields: dict, excel_df: pd.DataFrame):
        """
        Args:
            pdf_fields: dict of field_name -> value (from PLNExtractor.extract_flat())
            excel_df: DataFrame from the 'Data Pelanggan' sheet of the template
        """
        self.pdf_fields = pdf_fields
        self.excel_df = excel_df

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

        return {
            "matched_row": matched_row_idx,
            "results": results,
            "summary": summary,
        }

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
        """Find the Excel row that best matches the PDF by ID or meter number."""
        pdf_id = _normalize_numeric(self.pdf_fields.get("id_pelanggan", ""))
        pdf_meter = _normalize(self.pdf_fields.get("nomor_meter", ""))

        best_idx = None
        best_score = 0

        for idx in range(len(self.excel_df)):
            row = self.excel_df.iloc[idx]
            score = 0

            # Match by ID Pelanggan
            excel_id = _normalize_numeric(row.get("ID Pelanggan", ""))
            if pdf_id and excel_id and pdf_id == excel_id:
                score += 10

            # Match by Nomor Meter
            excel_meter = _normalize(row.get("Nomor Meter", ""))
            if pdf_meter and excel_meter and pdf_meter == excel_meter:
                score += 5

            # Match by name (partial)
            pdf_name = _normalize(self.pdf_fields.get("nama_pelanggan", ""))
            excel_name = _normalize(row.get("Nama Pelanggan", ""))
            if pdf_name and excel_name and _fuzzy_contains(pdf_name, excel_name):
                score += 3

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

        # String comparison
        pdf_norm = _normalize(pdf_value)
        excel_norm = _normalize(excel_value)

        if pdf_norm == excel_norm:
            result["match_status"] = "MATCH"
            result["notes"] = "Exact match"
        elif _fuzzy_contains(pdf_norm, excel_norm):
            result["match_status"] = "MATCH"
            result["notes"] = "Partial/contains match"
        else:
            result["match_status"] = "MISMATCH"
            result["notes"] = f"PDF='{pdf_value}' vs Excel='{excel_value}'"

        return result

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
