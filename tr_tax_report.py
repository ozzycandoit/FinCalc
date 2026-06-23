"""
tr_tax_report.py - renders the GiB "Hazir Beyan" PDF outputs:
  1) Hazir Beyan Sistemi Ozeti  (declaration summary: 3.C dividends, 3.D gains)
  2) Yatirim Islemleri Detayli Raporu  (every processed transaction)
Turkish characters are supported via an embedded DejaVuSans font.
"""

import io
import os

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, PageBreak)

NAVY = colors.HexColor("#1B2A4A")
TEAL = colors.HexColor("#138D9C")
LIGHT = colors.HexColor("#EAF3E6")
GREY = colors.HexColor("#F2F2F2")

# Core functions and implementation would go here
