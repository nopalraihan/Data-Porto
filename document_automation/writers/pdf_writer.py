"""
PDFWriter - Generate formatted PDF reports from DataFrames and text.
"""

import os
import unicodedata
from datetime import datetime

import pandas as pd
from fpdf import FPDF


class PDFWriter:
    """Generate professional PDF reports with tables, text, and summaries."""

    # Color palette (R, G, B)
    PRIMARY = (47, 84, 150)
    HEADER_BG = (47, 84, 150)
    HEADER_TEXT = (255, 255, 255)
    ALT_ROW = (242, 242, 242)
    WHITE = (255, 255, 255)
    BLACK = (0, 0, 0)
    GRAY = (100, 100, 100)
    LIGHT_GRAY = (200, 200, 200)

    def __init__(self, output_path: str, orientation: str = "P"):
        self.output_path = os.path.abspath(output_path)
        self.pdf = FPDF(orientation=orientation, unit="mm", format="A4")
        self.pdf.set_auto_page_break(auto=True, margin=20)
        self.pdf.add_page()
        self._page_width = self.pdf.w - 2 * self.pdf.l_margin

    @staticmethod
    def _sanitize_text(text: str) -> str:
        """Replace Unicode characters unsupported by latin-1 with ASCII equivalents."""
        # Normalize unicode ligatures and special chars (e.g. ﬁ -> fi)
        text = unicodedata.normalize("NFKD", text)
        # Encode to latin-1, replacing remaining unsupported chars
        return text.encode("latin-1", errors="replace").decode("latin-1")

    def add_title_page(self, title: str, subtitle: str = None, author: str = "Document Automation") -> None:
        """Add a centered title page."""
        self.pdf.set_y(80)
        self.pdf.set_font("Helvetica", "B", 28)
        self.pdf.set_text_color(*self.PRIMARY)
        self.pdf.cell(0, 15, self._sanitize_text(title), align="C", new_x="LMARGIN", new_y="NEXT")

        if subtitle:
            self.pdf.set_font("Helvetica", "", 14)
            self.pdf.set_text_color(*self.GRAY)
            self.pdf.cell(0, 10, self._sanitize_text(subtitle), align="C", new_x="LMARGIN", new_y="NEXT")

        self.pdf.ln(20)

        # Horizontal rule
        self.pdf.set_draw_color(*self.PRIMARY)
        self.pdf.set_line_width(0.5)
        x_start = self.pdf.l_margin + self._page_width * 0.2
        x_end = self.pdf.l_margin + self._page_width * 0.8
        self.pdf.line(x_start, self.pdf.get_y(), x_end, self.pdf.get_y())

        self.pdf.ln(10)
        self.pdf.set_font("Helvetica", "", 10)
        self.pdf.set_text_color(*self.GRAY)
        self.pdf.cell(0, 8, f"Author: {author}", align="C", new_x="LMARGIN", new_y="NEXT")
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.pdf.cell(0, 8, f"Generated: {timestamp}", align="C", new_x="LMARGIN", new_y="NEXT")

    def add_heading(self, text: str, level: int = 1) -> None:
        """Add a section heading (level 1-3)."""
        sizes = {1: 16, 2: 13, 3: 11}
        font_size = sizes.get(level, 11)

        self.pdf.ln(6)
        self.pdf.set_font("Helvetica", "B", font_size)
        self.pdf.set_text_color(*self.PRIMARY)
        self.pdf.cell(0, 10, self._sanitize_text(text), new_x="LMARGIN", new_y="NEXT")

        if level == 1:
            self.pdf.set_draw_color(*self.PRIMARY)
            self.pdf.set_line_width(0.3)
            self.pdf.line(self.pdf.l_margin, self.pdf.get_y(), self.pdf.l_margin + self._page_width, self.pdf.get_y())
            self.pdf.ln(3)

    def add_paragraph(self, text: str) -> None:
        """Add a paragraph of body text."""
        self.pdf.set_font("Helvetica", "", 10)
        self.pdf.set_text_color(*self.BLACK)
        self.pdf.multi_cell(0, 6, self._sanitize_text(text))
        self.pdf.ln(3)

    def add_key_value_section(self, items: list[tuple[str, str]]) -> None:
        """Add a section of key-value pairs (e.g., overview stats)."""
        for key, value in items:
            self.pdf.set_font("Helvetica", "B", 10)
            self.pdf.set_text_color(*self.BLACK)
            self.pdf.cell(60, 7, self._sanitize_text(f"{key}:"), new_x="END")
            self.pdf.set_font("Helvetica", "", 10)
            self.pdf.cell(0, 7, self._sanitize_text(str(value)), new_x="LMARGIN", new_y="NEXT")
        self.pdf.ln(3)

    def add_dataframe_table(self, df: pd.DataFrame, title: str = None, max_rows: int = 50) -> None:
        """Render a DataFrame as a formatted table in the PDF."""
        if title:
            self.add_heading(title, level=2)

        display_df = df.head(max_rows)
        truncated = len(df) > max_rows

        n_cols = len(display_df.columns)
        col_width = self._page_width / n_cols
        # Cap column width at reasonable bounds
        col_width = min(col_width, 60)
        row_height = 7

        # Check if we need landscape or smaller font
        total_width = col_width * n_cols
        font_size = 8 if total_width > self._page_width else 9

        # Recalculate to fit
        col_width = self._page_width / n_cols

        # Header row
        self.pdf.set_font("Helvetica", "B", font_size)
        self.pdf.set_fill_color(*self.HEADER_BG)
        self.pdf.set_text_color(*self.HEADER_TEXT)
        for col_name in display_df.columns:
            text = self._sanitize_text(str(col_name)[:20])
            self.pdf.cell(col_width, row_height, text, border=1, fill=True, align="C", new_x="END")
        self.pdf.ln()

        # Data rows
        self.pdf.set_font("Helvetica", "", font_size)
        self.pdf.set_text_color(*self.BLACK)

        for row_idx, (_, row) in enumerate(display_df.iterrows()):
            if row_idx % 2 == 1:
                self.pdf.set_fill_color(*self.ALT_ROW)
                fill = True
            else:
                self.pdf.set_fill_color(*self.WHITE)
                fill = True

            for value in row:
                if hasattr(value, "item"):
                    value = value.item()
                if pd.isna(value):
                    text = ""
                elif isinstance(value, float):
                    text = f"{value:,.2f}"
                else:
                    text = str(value)[:25]
                self.pdf.cell(col_width, row_height, self._sanitize_text(text), border=1, fill=fill, align="C", new_x="END")
            self.pdf.ln()

        if truncated:
            self.pdf.ln(2)
            self.pdf.set_font("Helvetica", "I", 8)
            self.pdf.set_text_color(*self.GRAY)
            self.pdf.cell(0, 5, f"Showing {max_rows} of {len(df)} rows.", new_x="LMARGIN", new_y="NEXT")
        self.pdf.ln(5)

    def add_page_break(self) -> None:
        """Insert a page break."""
        self.pdf.add_page()

    def _add_footer(self) -> None:
        """Add page numbers to all pages (called before save)."""
        total = self.pdf.pages_count
        for i in range(1, total + 1):
            self.pdf.page = i
            self.pdf.set_y(-15)
            self.pdf.set_font("Helvetica", "I", 8)
            self.pdf.set_text_color(*self.GRAY)
            self.pdf.cell(0, 10, f"Page {i} of {total}", align="C")

    def save(self) -> str:
        """Save the PDF and return the output path."""
        os.makedirs(os.path.dirname(self.output_path), exist_ok=True)
        self._add_footer()
        self.pdf.output(self.output_path)
        return self.output_path
