import datetime
import functools
import logging
import pathlib
from typing import Optional, Callable
import polars as pl
import pandas as pd
import math

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape, portrait
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.pdfgen import canvas

# Optional Arabic shaping dependencies
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    HAS_ARABIC = True
except ImportError:
    HAS_ARABIC = False

FONT_NAME = "Amiri-Bold"
FONT_PATH = pathlib.Path(__file__).parent / "amiri-bold.ttf"

# F-04: Guard font registration — track whether Amiri was actually registered
_AMIRI_REGISTERED = False
if FONT_PATH.exists():
    try:
        pdfmetrics.registerFont(TTFont(FONT_NAME, str(FONT_PATH)))
        FALLBACK_FONT = FONT_NAME
        _AMIRI_REGISTERED = True
    except Exception as _font_err:
        logging.warning(f"Impossible d'enregistrer la police Amiri : {_font_err}")
        FALLBACK_FONT = "Helvetica"
else:
    FALLBACK_FONT = "Helvetica"


@functools.lru_cache(maxsize=4096)
def _reshape_arabic(text: str) -> str:
    """F-03 + F-07: Reshape Arabic text for RTL rendering. Cached for performance.
    Returns the original string unchanged if shaping libraries are absent.
    Callers that need RTL must check HAS_ARABIC before trusting the result.
    """
    if not HAS_ARABIC or not text:
        return text
    reshaped_text = arabic_reshaper.reshape(str(text))
    bidi_text = get_display(reshaped_text)
    return bidi_text


class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_number(num_pages)
            super().showPage()
        super().save()

    def draw_page_number(self, page_count):
        self.setFont("Helvetica", 9)
        self.drawRightString(
            self._pagesize[0] - 1.5 * cm,
            2.0 * cm - 1 * cm,
            f"{self._pageNumber}/{page_count}"
        )



class PDFReportTemplate(BaseDocTemplate):

    def __init__(self, filename, title: str, subtitle: str, is_rtl: bool = False, **kw):
        super().__init__(filename, **kw)

        self.title = title
        self.subtitle = subtitle
        self._is_rtl = is_rtl

        frame = Frame(
            self.leftMargin,
            self.bottomMargin,
            self.width,
            self.height,
            id="normal"
        )

        template = PageTemplate(
            id="report_page",
            frames=frame,
            onPage=self._header_footer
        )

        self.addPageTemplates([template])

    def _header_footer(self, canvas, doc):

        canvas.saveState()

        # F-06: Warn when Amiri is absent but content may contain Arabic
        active_font = FALLBACK_FONT
        if self._is_rtl and not _AMIRI_REGISTERED:
            logging.warning(
                "amiri-bold.ttf introuvable : le titre/sous-titre arabe peut s'afficher incorrectement "
                "(glyphes vides avec Helvetica). Placez amiri-bold.ttf dans le dossier de l'application."
            )

        canvas.setFont(active_font, 16)

        safe_title = _reshape_arabic(str(self.title)) if self.title else ""

        canvas.drawCentredString(
            doc.pagesize[0] / 2.0,
            doc.pagesize[1] - doc.topMargin + 1 * cm,
            safe_title
        )

        if self.subtitle:
            canvas.setFont(active_font, 10)
            canvas.setFillColor(colors.dimgrey)

            safe_subtitle = _reshape_arabic(str(self.subtitle))

            canvas.drawCentredString(
                doc.pagesize[0] / 2.0,
                doc.pagesize[1] - doc.topMargin + 0.3 * cm,
                safe_subtitle
            )

            canvas.setFillColor(colors.black)

        canvas.setStrokeColor(colors.lightgrey)
        canvas.setLineWidth(1)

        canvas.line(
            doc.leftMargin,
            doc.pagesize[1] - doc.topMargin,
            doc.leftMargin + doc.width,
            doc.pagesize[1] - doc.topMargin
        )

        canvas.restoreState()


def generate_report(
    df,
    output_path: str,
    title: str = "Rapport de Données",
    subtitle: str = "",
    orientation: str = "Portrait",
    direction: str = "Français (LTR)",
    progress_callback: Optional[Callable] = None
):
    """
    Generate PDF report from Excel data.
    Accepts both Pandas and Polars DataFrames.
    """

    # -----------------------------
    # Convert Pandas -> Polars
    # -----------------------------
    if isinstance(df, pd.DataFrame):
        df = pl.from_pandas(df)

    # Fill null values
    df = df.fill_null("")

    # -----------------------------
    # RTL detection
    # -----------------------------
    direction_lower = direction.lower()

    is_rtl = any(kw in direction_lower for kw in [
        "ar",
        "arabic",
        "arabe",
        "rtl",
        "arab"
    ])

    # F-03: Raise early if RTL requested but shaping libraries are missing
    if is_rtl and not HAS_ARABIC:
        raise RuntimeError(
            "La génération de PDF en arabe (RTL) nécessite les bibliothèques "
            "'arabic-reshaper' et 'python-bidi'.\n"
            "Installez-les avec : pip install arabic-reshaper python-bidi"
        )

    # F-04: Raise early if RTL requested but Amiri font is missing
    if is_rtl and not _AMIRI_REGISTERED:
        raise RuntimeError(
            "La génération de PDF en arabe (RTL) nécessite la police 'amiri-bold.ttf'.\n"
            f"Placez le fichier dans : {FONT_PATH}"
        )

    # -----------------------------
    # Page setup
    # -----------------------------
    is_landscape = orientation.lower() == "paysage"

    page_size = landscape(A4) if is_landscape else portrait(A4)

    left_margin = 1.5 * cm
    right_margin = 1.5 * cm
    top_margin = 2.5 * cm
    bottom_margin = 2.0 * cm

    doc = PDFReportTemplate(
        output_path,
        title=title,
        subtitle=subtitle,
        is_rtl=is_rtl,
        pagesize=page_size,
        leftMargin=left_margin,
        rightMargin=right_margin,
        topMargin=top_margin,
        bottomMargin=bottom_margin
    )

    # -----------------------------
    # Extract columns
    # -----------------------------
    # F-01: Do NOT reverse columns globally — this scrambles mixed-language sheets.
    # RTL visual order is achieved through paragraph alignment (alignment=2 = RIGHT)
    # and TableStyle ALIGN directives below, not by reversing the column list.
    columns = list(df.columns)

    styles = getSampleStyleSheet()

    header_style = ParagraphStyle(
        "HeaderStyle",
        parent=styles["Normal"],
        fontName=FALLBACK_FONT,
        fontSize=10,
        textColor=colors.whitesmoke,
        alignment=2 if is_rtl else 1   # 2=RIGHT for RTL, 1=CENTER for LTR
    )

    cell_style = ParagraphStyle(
        "CellStyle",
        parent=styles["Normal"],
        fontName=FALLBACK_FONT,
        fontSize=8,
        alignment=2 if is_rtl else 0   # 2=RIGHT for RTL, 0=LEFT for LTR
    )

    available_width = page_size[0] - left_margin - right_margin
    col_count = max(len(columns), 1)
    col_width = available_width / col_count

    # F-02: Always cast to str before passing to _reshape_arabic (column names
    # may be integers or other non-str types in Polars).
    header_row = [
        Paragraph(
            _reshape_arabic(str(col)) if is_rtl else str(col),
            header_style
        )
        for col in columns
    ]

    elements = []
    total_rows = len(df)
    
    # Fix ISSUE-003: Chunk data so ReportLab doesn't hold millions of Paragraphs in memory
    CHUNK_SIZE = 500
    
    for chunk_start in range(0, max(total_rows, 1), CHUNK_SIZE):
        chunk_df = df.slice(chunk_start, CHUNK_SIZE) if total_rows > 0 else pl.DataFrame()
        rows = chunk_df.to_dicts()

        table_data = [header_row]

        for i, row_dict in enumerate(rows):
            global_row_idx = chunk_start + i
            row_cells = []

            for col in columns:
                val = row_dict.get(col, "")
                if val is None or (isinstance(val, float) and math.isnan(val)):
                    v_str = ""
                else:
                    v_str = str(val)

                if is_rtl:
                    v_str = _reshape_arabic(v_str)

                row_cells.append(Paragraph(v_str, cell_style))

            table_data.append(row_cells)

            # progress callback
            if progress_callback:
                progress_callback(global_row_idx + 1, total_rows)

        if not rows and chunk_start == 0:
            elements.append(Paragraph("Aucune donnée", cell_style))
            continue

        t = Table(
            table_data,
            colWidths=[col_width] * col_count,
            repeatRows=1
        )

        t_style = TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#313244")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("ALIGN", (0, 0), (-1, 0), "RIGHT" if is_rtl else "CENTER"),
            # F-05: Align data cells RIGHT for RTL so Arabic text sits against the right edge
            ("ALIGN", (0, 1), (-1, -1), "RIGHT" if is_rtl else "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
            ("TOPPADDING", (0, 0), (-1, 0), 8),
            # F-05: Swap left/right padding for RTL to give breathing room on the right side
            ("LEFTPADDING",  (0, 0), (-1, -1), 4 if is_rtl else 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8 if is_rtl else 6),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cdd6f4")),
            ("FONTNAME", (0, 1), (-1, -1), FALLBACK_FONT),
            ("FONTSIZE", (0, 1), (-1, -1), 8),
        ])

        # zebra rows
        for i in range(1, len(table_data)):
            global_idx = chunk_start + (i - 1)
            if global_idx % 2 == 0:
                t_style.add("BACKGROUND", (0, i), (-1, i), colors.HexColor("#e6e6e6"))
            else:
                t_style.add("BACKGROUND", (0, i), (-1, i), colors.white)

        t.setStyle(t_style)
        elements.append(t)
        
        # Add a small buffer between chunks visually if desired. The repeatRows=1 ensures header prints on cuts 
        # By separating chunks into tables, ReportLab handles page breaks natively without OOM.

    doc.build(elements, canvasmaker=NumberedCanvas)