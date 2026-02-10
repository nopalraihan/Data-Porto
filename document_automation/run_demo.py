#!/usr/bin/env python3
"""
Demo script - Demonstrates the full document automation pipeline.

Reads a CSV file, processes the data, and generates:
  1. A formatted Excel workbook with multiple sheets and charts
  2. A professional PDF report with tables and summaries

Usage:
    python run_demo.py
    python run_demo.py --input path/to/file.csv
    python run_demo.py --input data.xlsx --no-pdf
"""

import argparse
import os
import sys

# Add parent directory to path for module resolution
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from document_automation.pipeline import Pipeline


def run_csv_demo(input_path: str, output_dir: str):
    """Run the pipeline on a CSV/Excel file with sample transformations."""
    print(f"\n{'='*60}")
    print(f"  Document Automation Pipeline")
    print(f"{'='*60}")
    print(f"  Input:  {input_path}")
    print(f"  Output: {output_dir}")
    print(f"{'='*60}\n")

    pipeline = Pipeline(input_path, output_dir=output_dir)

    # Show file summary
    from document_automation.readers.document_reader import DocumentReader
    reader = DocumentReader(input_path)
    print("[1/4] Reading document...")
    print(reader.summary())
    print()

    # Run pipeline with sample filters and grouping
    print("[2/4] Processing data...")
    result = pipeline.run(
        title="Employee Analytics Report",
        filters=[
            {"column": "salary", "operator": ">=", "value": 60000},
        ],
        group_by={
            "group_col": "department",
            "agg_col": "salary",
            "agg_func": "mean",
        },
        top_n={
            "column": "salary",
            "n": 10,
            "ascending": False,
        },
    )

    # Report results
    if "excel" in result:
        print(f"[3/4] Excel report saved: {result['excel']}")
    else:
        print("[3/4] Excel generation skipped.")

    if "pdf" in result:
        print(f"[4/4] PDF report saved:   {result['pdf']}")
    else:
        print("[4/4] PDF generation skipped.")

    print(f"\nSummary: {result['summary']}")
    print("\nDone.")


def run_pdf_demo(input_path: str, output_dir: str):
    """Run the pipeline on a PDF file to extract and re-export its content."""
    print(f"\n{'='*60}")
    print(f"  PDF Text Extraction Pipeline")
    print(f"{'='*60}")
    print(f"  Input:  {input_path}")
    print(f"  Output: {output_dir}")
    print(f"{'='*60}\n")

    pipeline = Pipeline(input_path, output_dir=output_dir)

    from document_automation.readers.document_reader import DocumentReader
    reader = DocumentReader(input_path)
    print("[1/2] Reading PDF...")
    print(reader.summary())
    print()

    print("[2/2] Generating extracted report...")
    result = pipeline.run(title="Extracted PDF Content")

    if "pdf" in result:
        print(f"  PDF report saved: {result['pdf']}")
    print(f"\nSummary: {result['summary']}")
    print("\nDone.")


def main():
    parser = argparse.ArgumentParser(description="Document Automation Pipeline Demo")
    parser.add_argument(
        "--input", "-i",
        help="Path to input file (CSV, Excel, or PDF). Defaults to sample data.",
        default=None,
    )
    parser.add_argument(
        "--output", "-o",
        help="Output directory. Defaults to ./output/",
        default=None,
    )
    args = parser.parse_args()

    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = args.output or os.path.join(script_dir, "output")

    if args.input:
        input_path = os.path.abspath(args.input)
    else:
        input_path = os.path.join(script_dir, "sample_data", "employees.csv")

    if not os.path.isfile(input_path):
        print(f"Error: File not found: {input_path}")
        sys.exit(1)

    ext = os.path.splitext(input_path)[1].lower()
    if ext == ".pdf":
        run_pdf_demo(input_path, output_dir)
    else:
        run_csv_demo(input_path, output_dir)


if __name__ == "__main__":
    main()
