"""
PLNExtractor - Extract structured fields from PLN (electricity) PDF documents.

Handles various PLN document formats including:
- Surat Perjanjian Jual Beli Tenaga Listrik (SPJBTL)
- Rekening/Tagihan Listrik
- Data Pelanggan / Instalasi
"""

import re
import os
from datetime import datetime
from PyPDF2 import PdfReader


class PLNExtractor:
    """Extract key fields from PLN PDF documents into a structured dict."""

    # Regex patterns for common PLN document fields
    PATTERNS = {
        "id_pelanggan": [
            r"(?:ID\s*Pelanggan|IDPEL|No\.?\s*Pelanggan|Nomor\s*Pelanggan)\s*[:\-]?\s*(\d[\d\s\.\-]{6,})",
            r"(\d{12})",  # PLN customer IDs are typically 12 digits
        ],
        "nama_pelanggan": [
            r"(?:Nama\s*Pelanggan|Nama\s*Pemilik|Nama|Pelanggan|Atas\s*Nama)\s*[:\-]?\s*([A-Z][A-Za-z\s\.\,\']{2,50})",
        ],
        "alamat": [
            r"(?:Alamat|Alamat\s*Pelanggan|Alamat\s*Rumah)\s*[:\-]?\s*(.{10,100})",
        ],
        "tarif_daya": [
            r"(?:Tarif\s*/?\s*Daya|Tarif|Daya\s*Tersambung|Gol\.?\s*Tarif)\s*[:\-]?\s*([\w\d]+\s*/?\s*[\d\.]+\s*(?:VA|W|KVA|kVA)?)",
            r"(R[- ]?\d[A-Z]?\s*/?\s*[\d\.]+\s*(?:VA|W|KVA)?)",  # e.g., R1/1300 VA
        ],
        "nomor_meter": [
            r"(?:No\.?\s*Meter|Nomor\s*Meter|No\.?\s*APP|Meter\s*No)\s*[:\-]?\s*(\w[\w\d\-\.]{3,20})",
        ],
        "nomor_kwh": [
            r"(?:No\.?\s*KWH|kWh\s*Meter|Nomor\s*kWh)\s*[:\-]?\s*(\w[\w\d\-\.]{3,20})",
        ],
        "stand_meter_awal": [
            r"(?:Stand\s*(?:Meter\s*)?Awal|Meter\s*Awal|LWBP\s*Awal|Stand\s*Awal)\s*[:\-]?\s*([\d\.\,]+)",
        ],
        "stand_meter_akhir": [
            r"(?:Stand\s*(?:Meter\s*)?Akhir|Meter\s*Akhir|LWBP\s*Akhir|Stand\s*Akhir)\s*[:\-]?\s*([\d\.\,]+)",
        ],
        "pemakaian_kwh": [
            r"(?:Pemakaian|Jumlah\s*(?:Pemakaian|kWh)|Total\s*(?:Pemakaian|kWh)|kWh\s*Pakai)\s*[:\-]?\s*([\d\.\,]+)\s*(?:kWh|kwh)?",
        ],
        "biaya_listrik": [
            r"(?:Total\s*(?:Tagihan|Bayar|Rekening)|Jumlah\s*(?:Tagihan|Bayar)|Biaya\s*(?:Listrik|Total)|RP\s*Tag)\s*[:\-]?\s*(?:Rp\.?\s*)?([\d\.\,]+)",
        ],
        "periode": [
            r"(?:Periode|Bulan|Bln|Periode\s*Rekening)\s*[:\-]?\s*(\w+\s*\d{4}|\d{2}\s*/?\-?\s*\d{4})",
        ],
        "nomor_referensi": [
            r"(?:No\.?\s*(?:Ref|Referensi|Agenda)|Ref\.?\s*No)\s*[:\-]?\s*([\w\d\-\/\.]+)",
            r"JTX\d+",  # JTX-style reference numbers
        ],
    }

    def __init__(self, pdf_path: str):
        self.pdf_path = os.path.abspath(pdf_path)
        if not os.path.isfile(self.pdf_path):
            raise FileNotFoundError(f"PDF not found: {self.pdf_path}")
        self._raw_text = None
        self._pages = None

    def _read_pdf(self):
        """Read and cache PDF text content."""
        if self._raw_text is not None:
            return

        reader = PdfReader(self.pdf_path)
        self._pages = []
        all_text = []

        for i, page in enumerate(reader.pages):
            text = page.extract_text() or ""
            self._pages.append({"page_number": i + 1, "text": text})
            all_text.append(text)

        self._raw_text = "\n".join(all_text)

    def extract(self) -> dict:
        """
        Extract all recognized fields from the PDF.

        Returns:
            dict with keys:
                - 'fields': dict of field_name -> {'value': str, 'confidence': str, 'page': int}
                - 'metadata': dict with file info
                - 'raw_text': full extracted text
        """
        self._read_pdf()

        fields = {}
        for field_name, patterns in self.PATTERNS.items():
            match_result = self._find_field(field_name, patterns)
            if match_result:
                fields[field_name] = match_result

        metadata = {
            "file_name": os.path.basename(self.pdf_path),
            "file_path": self.pdf_path,
            "file_size_bytes": os.path.getsize(self.pdf_path),
            "page_count": len(self._pages),
            "extracted_at": datetime.now().isoformat(),
            "fields_found": len(fields),
            "fields_missing": len(self.PATTERNS) - len(fields),
        }

        return {
            "fields": fields,
            "metadata": metadata,
            "raw_text": self._raw_text,
            "pages": self._pages,
        }

    def _find_field(self, field_name: str, patterns: list) -> dict | None:
        """Try each pattern and return the first match with page location."""
        for pattern in patterns:
            # Search page by page to get location
            for page in self._pages:
                match = re.search(pattern, page["text"], re.IGNORECASE | re.MULTILINE)
                if match:
                    value = match.group(1) if match.lastindex else match.group(0)
                    value = value.strip()
                    return {
                        "value": value,
                        "confidence": "high" if len(patterns) == 1 else "medium",
                        "page": page["page_number"],
                        "pattern_used": pattern,
                    }

            # Also try full text (cross-page matches)
            match = re.search(pattern, self._raw_text, re.IGNORECASE | re.MULTILINE)
            if match:
                value = match.group(1) if match.lastindex else match.group(0)
                value = value.strip()
                return {
                    "value": value,
                    "confidence": "low",
                    "page": None,
                    "pattern_used": pattern,
                }

        return None

    def extract_flat(self) -> dict:
        """Return a simple field_name -> value dict (for crosschecking)."""
        result = self.extract()
        return {k: v["value"] for k, v in result["fields"].items()}

    def summary(self) -> str:
        """Return a human-readable summary of extracted fields."""
        result = self.extract()
        lines = [
            f"PLN Document: {result['metadata']['file_name']}",
            f"Pages: {result['metadata']['page_count']}",
            f"Fields found: {result['metadata']['fields_found']}/{len(self.PATTERNS)}",
            "",
            "Extracted Fields:",
        ]
        for name, info in result["fields"].items():
            lines.append(f"  {name:25s} = {info['value']:30s} (confidence: {info['confidence']}, page: {info['page']})")

        missing = set(self.PATTERNS.keys()) - set(result["fields"].keys())
        if missing:
            lines.append("")
            lines.append("Missing Fields:")
            for name in sorted(missing):
                lines.append(f"  {name}")

        return "\n".join(lines)
