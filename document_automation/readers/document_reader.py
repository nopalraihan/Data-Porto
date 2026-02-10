"""
DocumentReader - Unified reader for CSV, Excel, and PDF files.
"""

import os
import pandas as pd
from PyPDF2 import PdfReader


class DocumentReader:
    """Reads data from CSV, Excel (.xlsx/.xls), and PDF files."""

    SUPPORTED_EXTENSIONS = {".csv", ".xlsx", ".xls", ".pdf"}

    def __init__(self, file_path: str):
        self.file_path = os.path.abspath(file_path)
        self.extension = os.path.splitext(self.file_path)[1].lower()
        self._validate()

    def _validate(self):
        if not os.path.isfile(self.file_path):
            raise FileNotFoundError(f"File not found: {self.file_path}")
        if self.extension not in self.SUPPORTED_EXTENSIONS:
            raise ValueError(
                f"Unsupported file type '{self.extension}'. "
                f"Supported: {', '.join(sorted(self.SUPPORTED_EXTENSIONS))}"
            )

    def read(self, **kwargs) -> dict:
        """
        Read the file and return a dict with keys:
          - 'type': 'tabular' or 'text'
          - 'data': pd.DataFrame (tabular) or str (text)
          - 'metadata': dict with file info
        """
        metadata = {
            "file_name": os.path.basename(self.file_path),
            "file_path": self.file_path,
            "file_size_bytes": os.path.getsize(self.file_path),
            "format": self.extension,
        }

        if self.extension == ".csv":
            return self._read_csv(metadata, **kwargs)
        elif self.extension in (".xlsx", ".xls"):
            return self._read_excel(metadata, **kwargs)
        elif self.extension == ".pdf":
            return self._read_pdf(metadata, **kwargs)

    def _read_csv(self, metadata: dict, **kwargs) -> dict:
        df = pd.read_csv(self.file_path, **kwargs)
        metadata["row_count"] = len(df)
        metadata["column_count"] = len(df.columns)
        metadata["columns"] = list(df.columns)
        return {"type": "tabular", "data": df, "metadata": metadata}

    def _read_excel(self, metadata: dict, **kwargs) -> dict:
        sheet_name = kwargs.pop("sheet_name", None)
        if sheet_name is not None:
            df = pd.read_excel(self.file_path, sheet_name=sheet_name, **kwargs)
            metadata["sheet_name"] = sheet_name
            metadata["row_count"] = len(df)
            metadata["column_count"] = len(df.columns)
            metadata["columns"] = list(df.columns)
            return {"type": "tabular", "data": df, "metadata": metadata}

        # Read all sheets
        excel_file = pd.ExcelFile(self.file_path)
        sheets = {}
        for name in excel_file.sheet_names:
            sheets[name] = pd.read_excel(self.file_path, sheet_name=name, **kwargs)

        metadata["sheet_names"] = excel_file.sheet_names
        metadata["sheet_count"] = len(excel_file.sheet_names)
        return {"type": "tabular", "data": sheets, "metadata": metadata}

    def _read_pdf(self, metadata: dict, **kwargs) -> dict:
        reader = PdfReader(self.file_path)
        pages = []
        full_text = []

        for i, page in enumerate(reader.pages):
            text = page.extract_text() or ""
            pages.append({"page_number": i + 1, "text": text})
            full_text.append(text)

        metadata["page_count"] = len(reader.pages)
        return {
            "type": "text",
            "data": "\n\n".join(full_text),
            "pages": pages,
            "metadata": metadata,
        }

    def summary(self) -> str:
        """Return a human-readable summary of the file."""
        result = self.read()
        meta = result["metadata"]
        lines = [
            f"File: {meta['file_name']}",
            f"Format: {meta['format']}",
            f"Size: {meta['file_size_bytes']:,} bytes",
        ]

        if result["type"] == "tabular":
            if isinstance(result["data"], dict):
                lines.append(f"Sheets: {meta['sheet_count']}")
                for name, df in result["data"].items():
                    lines.append(f"  - {name}: {len(df)} rows x {len(df.columns)} cols")
            else:
                lines.append(f"Rows: {meta['row_count']}")
                lines.append(f"Columns: {meta['column_count']}")
                lines.append(f"Column names: {', '.join(meta['columns'])}")
        else:
            lines.append(f"Pages: {meta['page_count']}")
            char_count = len(result["data"])
            lines.append(f"Characters: {char_count:,}")

        return "\n".join(lines)
