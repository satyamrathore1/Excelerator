"""
pdf_builder.py
--------------
Fast, manual-pagination PDF writer for report tables.

Deliberately avoids reportlab's Table flowable for large datasets - Table's
automatic page-splitting measured ~5x slower at realistic scale (16s vs 3.4s
for a single 15,653-row table), which matters here since Fusion Inventory
(~15.6k rows), MHS Snapshot (~17.8k rows), and SCE Snapshot (~9.9k rows) are
all in that range and this runs synchronously during "Generate Reports".

Every cell is drawn with a full grid border (matches the xlsx's bordered
tables). Column width is a simple equal split, so very wide tables (Fusion
Inventory's 18 columns) end up with narrow columns and truncated cell text.
The xlsx remains the fully-readable, primary format for those reports - the
PDF is best treated as an archival companion, not something read
column-by-column on screen.
"""

from __future__ import annotations

from reportlab.lib.pagesizes import A4, landscape as _landscape_pagesize
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas


class PdfWriter:
    def __init__(self, pdf_path: str, landscape: bool = True):
        pagesize = _landscape_pagesize(A4) if landscape else A4
        self.page_w, self.page_h = pagesize
        self.margin = 1.2 * cm
        self.row_height = 12
        self.canvas = canvas.Canvas(pdf_path, pagesize=pagesize)
        self.canvas.setLineWidth(0.4)
        self.y = self.page_h - self.margin

    def _new_page(self):
        self.canvas.showPage()
        self.canvas.setLineWidth(0.4)
        self.y = self.page_h - self.margin

    def force_page_break(self):
        """Public hook for callers that want a section to start on its own
        fresh page - e.g. a Summary section meant to stand alone as page 1."""
        self._new_page()

    def add_title(self, text: str):
        """A large title line at the top of the document (or current page)."""
        if self.y < self.margin + self.row_height * 3:
            self._new_page()
        self.canvas.setFont("Helvetica-Bold", 14)
        self.canvas.drawString(self.margin, self.y, text)
        self.y -= self.row_height * 2

    def _draw_row(self, values: list, col_width: float, max_chars: int, bold: bool = False):
        """Draws one row of bordered cells and advances self.y by one row height."""
        self.canvas.setFont("Helvetica-Bold" if bold else "Helvetica", 7)
        row_bottom = self.y - self.row_height
        for i, val in enumerate(values):
            x = self.margin + i * col_width
            self.canvas.rect(x, row_bottom, col_width, self.row_height, stroke=1, fill=0)
            text = "" if val is None else str(val)
            self.canvas.drawString(x + 2, row_bottom + 3, text[:max_chars])
        self.y = row_bottom

    def add_section(self, title: str | None, headers: list[str], rows: list[list]):
        """A titled, fully-bordered table: header row repeats at the top of every page it spans."""
        col_width = (self.page_w - 2 * self.margin) / max(len(headers), 1)
        max_chars = max(1, int(col_width / 4.0))  # ~4pt per character at 7pt font

        if title:
            if self.y < self.margin + self.row_height * 4:
                self._new_page()
            self.canvas.setFont("Helvetica-Bold", 11)
            self.canvas.drawString(self.margin, self.y, title)
            self.y -= self.row_height * 1.5

        if self.y < self.margin + self.row_height * 2:
            self._new_page()
        self._draw_row(headers, col_width, max_chars, bold=True)

        if not rows:
            self.canvas.setFont("Helvetica", 7)
            self.canvas.drawString(self.margin, self.y - self.row_height, "(no rows)")
            self.y -= self.row_height

        for row in rows:
            if self.y < self.margin + self.row_height:
                self._new_page()
                self._draw_row(headers, col_width, max_chars, bold=True)
            self._draw_row(row, col_width, max_chars)

        self.y -= self.row_height  # gap after section

    def save(self):
        self.canvas.save()
