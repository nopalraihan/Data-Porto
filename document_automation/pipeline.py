"""
Pipeline - Orchestrates the document automation workflow:
  1. Read source documents (CSV, Excel, PDF)
  2. Process and transform data
  3. Export to formatted Excel workbook
  4. Export to professional PDF report
"""

import os
from datetime import datetime

import pandas as pd

from .readers.document_reader import DocumentReader
from .writers.excel_writer import ExcelWriter
from .writers.pdf_writer import PDFWriter
from .utils.data_processor import DataProcessor


class Pipeline:
    """
    End-to-end document automation pipeline.

    Usage:
        pipeline = Pipeline("input.csv", output_dir="output/")
        pipeline.run()
    """

    def __init__(self, input_path: str, output_dir: str = None):
        self.input_path = os.path.abspath(input_path)
        self.reader = DocumentReader(self.input_path)

        base_name = os.path.splitext(os.path.basename(self.input_path))[0]
        if output_dir is None:
            output_dir = os.path.join(os.path.dirname(self.input_path), "..", "output")
        self.output_dir = os.path.abspath(output_dir)
        os.makedirs(self.output_dir, exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.excel_path = os.path.join(self.output_dir, f"{base_name}_report_{timestamp}.xlsx")
        self.pdf_path = os.path.join(self.output_dir, f"{base_name}_report_{timestamp}.pdf")

    def run(
        self,
        excel: bool = True,
        pdf: bool = True,
        title: str = None,
        filters: list[dict] = None,
        group_by: dict = None,
        top_n: dict = None,
    ) -> dict:
        """
        Execute the full pipeline.

        Args:
            excel: Generate Excel output.
            pdf: Generate PDF output.
            title: Report title (auto-generated if None).
            filters: List of dicts with keys: column, operator, value.
            group_by: Dict with keys: group_col, agg_col, agg_func.
            top_n: Dict with keys: column, n, ascending.

        Returns:
            Dict with paths to generated files and processing summary.
        """
        result = self.reader.read()
        meta = result["metadata"]
        report_title = title or f"Report: {meta['file_name']}"

        # Handle PDF text input
        if result["type"] == "text":
            return self._process_text_document(result, report_title, pdf)

        # Handle tabular data
        df = result["data"]
        if isinstance(df, dict):
            # Multi-sheet Excel: use first sheet
            first_sheet = list(df.keys())[0]
            df = df[first_sheet]

        processor = DataProcessor(df)
        stats_df = processor.get_summary_statistics()
        column_info = processor.get_column_info()

        # Apply optional transformations
        filtered_df = None
        if filters:
            filtered_processor = processor
            for f in filters:
                filtered_processor = filtered_processor.filter_rows(
                    f["column"], f["operator"], f["value"]
                )
            filtered_df = filtered_processor.to_dataframe()

        grouped_df = None
        if group_by:
            grouped_df = processor.group_summary(
                group_by["group_col"],
                group_by["agg_col"],
                group_by.get("agg_func", "mean"),
            )

        top_n_df = None
        if top_n:
            top_n_df = processor.top_n(
                top_n["column"],
                top_n.get("n", 10),
                top_n.get("ascending", False),
            )

        outputs = {}

        if excel:
            outputs["excel"] = self._generate_excel(
                df, stats_df, column_info, report_title,
                filtered_df=filtered_df,
                grouped_df=grouped_df,
                top_n_df=top_n_df,
                group_by=group_by,
            )

        if pdf:
            outputs["pdf"] = self._generate_pdf(
                df, stats_df, column_info, meta, report_title,
                filtered_df=filtered_df,
                grouped_df=grouped_df,
                top_n_df=top_n_df,
                filters=filters,
                group_by=group_by,
                top_n=top_n,
            )

        outputs["summary"] = {
            "input_file": meta["file_name"],
            "rows": len(df),
            "columns": len(df.columns),
            "generated_at": datetime.now().isoformat(),
        }

        return outputs

    def _generate_excel(
        self, df, stats_df, column_info, title,
        filtered_df=None, grouped_df=None, top_n_df=None, group_by=None,
    ) -> str:
        writer = ExcelWriter(self.excel_path)

        # Summary sheet
        writer.add_summary_sheet(df, stats_df, column_info, sheet_name="Summary")

        # Full data sheet
        writer.add_dataframe_sheet(df, sheet_name="Full Data", title=title)

        # Filtered data
        if filtered_df is not None and not filtered_df.empty:
            writer.add_dataframe_sheet(filtered_df, sheet_name="Filtered Data", title="Filtered Results")

        # Grouped data with chart
        if grouped_df is not None and not grouped_df.empty:
            agg_col_name = [c for c in grouped_df.columns if c != group_by["group_col"]][0]
            writer.add_dataframe_sheet(
                grouped_df,
                sheet_name="Grouped Analysis",
                title="Grouped Analysis",
                include_chart=True,
                chart_col=agg_col_name,
                chart_label_col=group_by["group_col"],
            )

        # Top N
        if top_n_df is not None and not top_n_df.empty:
            writer.add_dataframe_sheet(top_n_df, sheet_name="Top Records", title="Top Records")

        return writer.save()

    def _generate_pdf(
        self, df, stats_df, column_info, meta, title,
        filtered_df=None, grouped_df=None, top_n_df=None,
        filters=None, group_by=None, top_n=None,
    ) -> str:
        # Use landscape for wide datasets
        orientation = "L" if len(df.columns) > 6 else "P"
        writer = PDFWriter(self.pdf_path, orientation=orientation)

        # Title page
        writer.add_title_page(title, subtitle=f"Source: {meta['file_name']}")

        # Overview section
        writer.add_page_break()
        writer.add_heading("Data Overview", level=1)
        writer.add_key_value_section([
            ("Source File", meta["file_name"]),
            ("File Format", meta["format"]),
            ("File Size", f"{meta['file_size_bytes']:,} bytes"),
            ("Total Rows", f"{len(df):,}"),
            ("Total Columns", str(len(df.columns))),
            ("Null Values", f"{int(df.isnull().sum().sum()):,}"),
        ])

        # Column info table
        writer.add_dataframe_table(column_info, title="Column Information")

        # Statistics
        if not stats_df.empty:
            writer.add_page_break()
            stats_reset = stats_df.reset_index()
            stats_reset.rename(columns={"index": "Statistic"}, inplace=True)
            writer.add_dataframe_table(stats_reset, title="Descriptive Statistics")

        # Data preview
        writer.add_page_break()
        writer.add_dataframe_table(df, title="Data Preview", max_rows=30)

        # Filtered results
        if filtered_df is not None and not filtered_df.empty:
            writer.add_page_break()
            filter_desc = ", ".join(
                f"{f['column']} {f['operator']} {f['value']}" for f in filters
            )
            writer.add_heading("Filtered Results", level=1)
            writer.add_paragraph(f"Filters applied: {filter_desc}")
            writer.add_paragraph(f"Matching rows: {len(filtered_df):,}")
            writer.add_dataframe_table(filtered_df, max_rows=30)

        # Grouped analysis
        if grouped_df is not None and not grouped_df.empty:
            writer.add_page_break()
            writer.add_heading("Grouped Analysis", level=1)
            writer.add_paragraph(
                f"Grouped by '{group_by['group_col']}', "
                f"aggregated '{group_by['agg_col']}' using {group_by.get('agg_func', 'mean')}."
            )
            writer.add_dataframe_table(grouped_df)

        # Top N
        if top_n_df is not None and not top_n_df.empty:
            writer.add_page_break()
            direction = "Bottom" if top_n.get("ascending") else "Top"
            writer.add_heading(f"{direction} {top_n.get('n', 10)} Records", level=1)
            writer.add_paragraph(f"Sorted by '{top_n['column']}'.")
            writer.add_dataframe_table(top_n_df)

        return writer.save()

    def _process_text_document(self, result: dict, title: str, generate_pdf: bool) -> dict:
        """Handle PDF text input - extract text and generate a PDF report."""
        meta = result["metadata"]
        outputs = {}

        if generate_pdf:
            writer = PDFWriter(self.pdf_path)
            writer.add_title_page(title, subtitle=f"Extracted from: {meta['file_name']}")

            writer.add_page_break()
            writer.add_heading("Document Information", level=1)
            writer.add_key_value_section([
                ("Source File", meta["file_name"]),
                ("Pages", str(meta["page_count"])),
                ("File Size", f"{meta['file_size_bytes']:,} bytes"),
            ])

            writer.add_heading("Extracted Content", level=1)
            for page in result["pages"]:
                writer.add_heading(f"Page {page['page_number']}", level=2)
                text = page["text"].strip()
                if text:
                    writer.add_paragraph(text)
                else:
                    writer.add_paragraph("[No extractable text on this page]")

            outputs["pdf"] = writer.save()

        outputs["summary"] = {
            "input_file": meta["file_name"],
            "pages": meta["page_count"],
            "characters": len(result["data"]),
            "generated_at": datetime.now().isoformat(),
        }

        return outputs
