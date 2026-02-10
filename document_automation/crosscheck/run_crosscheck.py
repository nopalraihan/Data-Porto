#!/usr/bin/env python3
"""
PLN Document Crosscheck Runner
===============================
Reads a PLN PDF document, compares it against an Excel template,
and generates a verification report (Excel + PDF) with match/mismatch details.

FLOW:
  1. Read PLN PDF document → extract fields (ID, nama, alamat, tarif, meter, etc.)
  2. Read Excel template (from HO PLN) → load expected data
  3. Crosscheck PDF fields vs Excel row
  4. If ALL match  → Document VERIFIED
  5. If mismatch   → Generate error log with exact fields that don't match

Usage:
    python run_crosscheck.py --pdf <path_to_pdf> --excel <path_to_template>
    python run_crosscheck.py --demo   (runs with sample data)
"""

import argparse
import os
import sys

# Add project root (Data-Porto/) to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from document_automation.crosscheck.pln_extractor import PLNExtractor
from document_automation.crosscheck.crosscheck_engine import CrosscheckEngine
from document_automation.crosscheck.report_generator import ReportGenerator
from document_automation.crosscheck.create_template import create_template

import pandas as pd


def print_banner(title: str):
    print(f"\n{'='*65}")
    print(f"  {title}")
    print(f"{'='*65}\n")


def print_field_result(r: dict):
    """Print a single field comparison result with color indicators."""
    status = r["match_status"]
    indicator = {"MATCH": "[OK]", "MISMATCH": "[!!]", "MISSING": "[??]"}[status]
    pdf_val = (r.get("pdf_value") or "—")[:35]
    xls_val = (r.get("excel_value") or "—")[:35]
    sim_info = ""
    if "similarity_score" in r:
        sim_info = f" (sim={r['similarity_score']:.2f})"
    print(f"  {indicator} {r['field_name']:25s} PDF: {pdf_val:35s} Excel: {xls_val}{sim_info}")
    if status == "MISMATCH":
        print(f"       >>> {r.get('notes', '')}")


def print_ml_results(result: dict):
    """Print ML analysis results (anomaly detection, similarity, confidence)."""
    print("\n  " + "=" * 60)
    print("  ML-POWERED ANALYSIS")
    print("  " + "=" * 60)

    # Text similarity
    if "ml_similarity" in result:
        print("\n  [TF-IDF Text Similarity]")
        for field, sim in result["ml_similarity"].items():
            print(f"    {field:25s} score={sim['score']:.4f}  ({sim['classification']})")

    # Anomaly detection
    if "ml_anomalies" in result:
        flags = result["ml_anomalies"]
        if flags:
            print(f"\n  [Anomaly Detection] {len(flags)} issue(s) found:")
            for f in flags:
                icon = "!!" if f["severity"] == "CRITICAL" else "**"
                print(f"    [{icon}] {f['severity']:8s} | {f['field']}: {f['message']}")
                print(f"         Expected: {f['expected']}  |  Actual: {f['actual']}")
        else:
            print("\n  [Anomaly Detection] No anomalies detected.")

    # Confidence score
    if "ml_confidence" in result:
        conf = result["ml_confidence"]
        print(f"\n  [Random Forest Confidence Scorer]")
        print(f"    Confidence Score:  {conf['confidence_score']}%")
        print(f"    Prediction:        {conf['prediction']}")
        print(f"    Risk Level:        {conf['risk_level']}")
        print(f"    Feature Weights:")
        for feat, weight in conf["feature_contributions"].items():
            bar = "#" * int(abs(weight) * 50)
            print(f"      {feat:25s} {weight:+.4f}  {bar}")

    print()


def run_crosscheck(pdf_path: str, excel_path: str, output_dir: str):
    """Full crosscheck pipeline."""
    print_banner("PLN DOCUMENT CROSSCHECK")
    print(f"  PDF:    {pdf_path}")
    print(f"  Excel:  {excel_path}")
    print(f"  Output: {output_dir}\n")

    # ── Step 1: Extract fields from PDF ──
    print("[1/4] Extracting fields from PDF document...")
    extractor = PLNExtractor(pdf_path)
    extraction = extractor.extract()
    pdf_fields = {k: v["value"] for k, v in extraction["fields"].items()}

    print(f"  Found {len(pdf_fields)} fields:")
    for k, v in pdf_fields.items():
        print(f"    {k:25s} = {v}")
    print()

    # ── Step 2: Load Excel template ──
    print("[2/4] Loading Excel template (HO PLN data)...")
    df = pd.read_excel(excel_path, sheet_name="Data Pelanggan", header=3)
    print(f"  Loaded {len(df)} rows from template")
    print()

    # ── Step 3: Run crosscheck ──
    print("[3/4] Running crosscheck comparison...")
    engine = CrosscheckEngine(pdf_fields, df)
    result = engine.run()

    summary = result["summary"]
    pct = summary["match_percentage"]

    # Print results
    print()
    if summary["matched_row_index"] is not None:
        print(f"  Matched to Excel Row #{summary['matched_row_number']}")
    else:
        print("  WARNING: No matching row found in Excel template!")
    print()

    print("  Field-by-Field Results:")
    print("  " + "-" * 100)
    for r in result["results"]:
        print_field_result(r)
    print("  " + "-" * 100)
    print()

    # ── Verdict ──
    if pct == 100:
        print(f"  >>> VERDICT: VERIFIED - All {summary['total_fields_checked']} fields match!")
        print("  >>> Document is VALID.\n")
    elif summary["matched_row_index"] is None:
        print("  >>> VERDICT: UNVERIFIED - No matching record found in Excel template.")
        print("  >>> Document cannot be verified.\n")
    else:
        print(f"  >>> VERDICT: MISMATCH DETECTED")
        print(f"  >>> {summary['total_match']}/{summary['total_fields_checked']} fields match ({pct}%)")
        print(f"  >>> {summary['total_mismatch']} MISMATCHES | {summary['total_missing']} MISSING\n")

        # Error log
        mismatches = [r for r in result["results"] if r["match_status"] == "MISMATCH"]
        if mismatches:
            print("  ERROR LOG:")
            for i, r in enumerate(mismatches, 1):
                print(f"    Error #{i}: {r['field_name']}")
                print(f"      PDF says:   {r.get('pdf_value', 'N/A')}")
                print(f"      Excel says: {r.get('excel_value', 'N/A')}")
                print(f"      Note: {r.get('notes', '')}")
                print()

    # ML analysis output
    print_ml_results(result)

    # ── Step 4: Generate reports ──
    print("[4/4] Generating output reports...")
    reporter = ReportGenerator(result, extraction["metadata"], output_dir)
    outputs = reporter.generate_all()
    print(f"  Excel report: {outputs['excel']}")
    print(f"  PDF report:   {outputs['pdf']}")
    print("\nDone.")

    return result


def run_demo(output_dir: str):
    """Run a demo with synthetic PDF-like data and the template."""
    print_banner("PLN CROSSCHECK - DEMO MODE")
    print("  Creating sample Excel template and simulating PDF extraction...\n")

    # Create template
    template_dir = os.path.join(os.path.dirname(__file__), "..", "sample_data")
    os.makedirs(template_dir, exist_ok=True)
    template_path = create_template(os.path.join(template_dir, "PLN_Crosscheck_Template.xlsx"))
    print(f"  Template created: {template_path}")

    # Simulate PDF extraction (as if we read the PDF "23.Dok PLN PENGGILINGANELOK_JTX476.pdf")
    # Data intentionally has 2 mismatches for demo purposes
    simulated_pdf_fields = {
        "id_pelanggan": "532100012345",
        "nama_pelanggan": "SUHARTO",
        "alamat": "JL. PENGGILINGAN ELOK NO.23 RT005/RW012, PENGGILINGAN, CAKUNG, JAKARTA TIMUR",
        "tarif_daya": "R1/1300 VA",
        "nomor_meter": "JTX476",
        "nomor_kwh": "85201234",
        "periode": "Januari 2026",
        "stand_meter_awal": "15230",
        "stand_meter_akhir": "15510",       # << MISMATCH: PDF says 15510, Excel says 15480
        "pemakaian_kwh": "280",              # << MISMATCH: PDF says 280, Excel says 250
        "biaya_listrik": "352500",
    }

    print(f"\n  Simulated PDF fields (with 2 intentional mismatches):")
    for k, v in simulated_pdf_fields.items():
        print(f"    {k:25s} = {v}")

    # Load template
    df = pd.read_excel(template_path, sheet_name="Data Pelanggan", header=3)
    print(f"\n  Template rows: {len(df)}")

    # Crosscheck
    print("\n  Running crosscheck...\n")
    engine = CrosscheckEngine(simulated_pdf_fields, df)
    result = engine.run()

    summary = result["summary"]
    pct = summary["match_percentage"]

    print("  Field-by-Field Results:")
    print("  " + "-" * 100)
    for r in result["results"]:
        print_field_result(r)
    print("  " + "-" * 100)

    if pct == 100:
        print(f"\n  >>> VERDICT: VERIFIED - Document is VALID.\n")
    else:
        print(f"\n  >>> VERDICT: MISMATCH DETECTED")
        print(f"  >>> {summary['total_match']}/{summary['total_fields_checked']} match ({pct}%)")
        print(f"  >>> {summary['total_mismatch']} MISMATCHES\n")

        mismatches = [r for r in result["results"] if r["match_status"] == "MISMATCH"]
        if mismatches:
            print("  ERROR LOG:")
            for i, r in enumerate(mismatches, 1):
                print(f"    Error #{i}: {r['field_name']}")
                print(f"      PDF says:   {r.get('pdf_value', 'N/A')}")
                print(f"      Excel says: {r.get('excel_value', 'N/A')}")
                print()

    # ML analysis
    print_ml_results(result)

    # Generate reports
    print("  Generating reports...")
    reporter = ReportGenerator(result, {
        "file_name": "23.Dok PLN PENGGILINGANELOK_JTX476.pdf",
        "page_count": 2,
    }, output_dir)
    outputs = reporter.generate_all()
    print(f"  Excel report: {outputs['excel']}")
    print(f"  PDF report:   {outputs['pdf']}")
    print("\nDemo complete.")

    return result


def main():
    parser = argparse.ArgumentParser(description="PLN Document Crosscheck")
    parser.add_argument("--pdf", "-p", help="Path to PLN PDF document")
    parser.add_argument("--excel", "-e", help="Path to PLN Excel template")
    parser.add_argument("--output", "-o", help="Output directory", default=None)
    parser.add_argument("--demo", action="store_true", help="Run demo with sample data")
    args = parser.parse_args()

    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = args.output or os.path.join(script_dir, "..", "output")

    if args.demo:
        run_demo(output_dir)
    elif args.pdf and args.excel:
        run_crosscheck(args.pdf, args.excel, output_dir)
    else:
        parser.print_help()
        print("\nExamples:")
        print("  python run_crosscheck.py --demo")
        print("  python run_crosscheck.py --pdf doc.pdf --excel template.xlsx")
        sys.exit(1)


if __name__ == "__main__":
    main()
