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

_FONT = "Helvetica"
_FONT_B = "Helvetica-Bold"


def _register_font():
    global _FONT, _FONT_B
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(here, "DejaVuSans.ttf"),
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    ]
    bold = [
        os.path.join(here, "DejaVuSans-Bold.ttf"),
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf",
    ]
    reg = next((p for p in candidates if os.path.exists(p)), None)
    bld = next((p for p in bold if os.path.exists(p)), None)
    if reg:
        pdfmetrics.registerFont(TTFont("DejaVu", reg))
        _FONT = "DejaVu"
        if bld:
            pdfmetrics.registerFont(TTFont("DejaVu-Bold", bld))
            _FONT_B = "DejaVu-Bold"
        else:
            _FONT_B = "DejaVu"


_register_font()


def _tl(x):
    try:
        return f"{float(x):,.2f} ₺"
    except (TypeError, ValueError):
        return "-"


def _styles():
    ss = getSampleStyleSheet()
    body = ParagraphStyle("body", parent=ss["Normal"], fontName=_FONT, fontSize=9, leading=12)
    h1 = ParagraphStyle("h1", parent=ss["Heading1"], fontName=_FONT_B, fontSize=15,
                        textColor=NAVY, spaceAfter=4)
    h2 = ParagraphStyle("h2", parent=ss["Heading2"], fontName=_FONT_B, fontSize=11,
                        textColor=TEAL, spaceBefore=10, spaceAfter=4)
    small = ParagraphStyle("small", parent=body, fontSize=8, textColor=colors.grey)
    return body, h1, h2, small


def _section_table(headers, rows, col_widths):
    data = [headers] + rows
    t = Table(data, colWidths=col_widths, repeatRows=1)
    style = [
        ("FONTNAME", (0, 0), (-1, 0), _FONT_B),
        ("FONTNAME", (0, 1), (-1, -1), _FONT),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("BACKGROUND", (0, 0), (-1, 0), TEAL),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 1), (-1, -1), LIGHT),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B7C7B0")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    t.setStyle(TableStyle(style))
    return t


def build_beyan_pdf(beyan, detail_rows, taxpayer_note=""):
    """beyan: dict from tr_tax_core.build_gib_beyan
       detail_rows: list from tr_tax_core.detailed_transactions
       Returns PDF bytes."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=15 * mm, bottomMargin=15 * mm,
                            leftMargin=14 * mm, rightMargin=14 * mm)
    body, h1, h2, small = _styles()
    story = []
    yr = beyan.get("tax_year", 2026)

    story.append(Paragraph(f"Hazır Beyan Sistemi Özeti — {yr}", h1))
    story.append(Paragraph("Tahmini özet. Resmi beyan değildir; mali müşavir onayı önerilir.", small))
    story.append(Spacer(1, 8))

    # 3.C dividends
    div = beyan.get("dividends")
    if div:
        story.append(Paragraph("3.C BEYAN EDİLECEK MENKUL SERMAYE İRADI GELİRLERİNİZ", h2))
        story.append(_section_table(
            ["İradın Türü", "Gayrisafi İrat", "İndirilecek Giderler", "Safi İrat", "Kesilen Gelir Vergisi"],
            [[div["code"], _tl(div["gayrisafi"]), _tl(div["indirilecek"]),
              _tl(div["safi"]), _tl(div["kesilen"])]],
            [70 * mm, 28 * mm, 30 * mm, 26 * mm, 28 * mm],
        ))
        story.append(Spacer(1, 8))

    # 3.D capital gains
    cap = beyan["capital_gains"]
    story.append(Paragraph("3.D BEYAN EDİLECEK DİĞER KAZANÇ VE İRAT GELİRLERİNİZ", h2))
    story.append(_section_table(
        ["Kazancın Türü", "Gayrisafi Tutar", "Gider / İndirim", "Safi Kazanç", "Kesilen Gelir Vergisi"],
        [[f"{cap['code']}\n{cap['label']}", _tl(cap["gayrisafi"]),
          _tl(cap["gider_indirim"]), _tl(cap["safi"]), _tl(cap["kesilen"])]],
        [70 * mm, 28 * mm, 28 * mm, 28 * mm, 28 * mm],
    ))
    story.append(Spacer(1, 10))

    # Tax box
    story.append(Paragraph("HESAPLANAN VERGİ", h2))
    story.append(_section_table(
        ["Matrah (vergiye esas)", f"Hesaplanan {yr} gelir vergisi", "1. Taksit (Mart)", "2. Taksit (Temmuz)"],
        [[_tl(beyan["tax_base"]), _tl(beyan["tax"]),
          _tl(beyan["instalment_1"]), _tl(beyan["instalment_2"])]],
        [42 * mm, 50 * mm, 35 * mm, 35 * mm],
    ))
    if beyan.get("status"):
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"Durum: {beyan['status']}", small))
    if taxpayer_note:
        story.append(Paragraph(taxpayer_note, small))

    # Detail report
    story.append(PageBreak())
    story.append(Paragraph(f"Yatırım İşlemleri Detaylı Raporu — {yr}", h1))
    story.append(Spacer(1, 6))
    if detail_rows:
        rows = []
        for t in detail_rows:
            d = t.get("date")
            dstr = d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else str(d)
            rows.append([
                str(t.get("asset", ""))[:22],
                dstr,
                str(t.get("type", "")),
                f"{t.get('qty', '') or ''}",
                f"{t.get('price', '') or ''}",
                _tl(t["tl_amount"]) if t.get("tl_amount") is not None else "-",
                _tl(t["gross"]) if t.get("gross") is not None else "-",
                _tl(t["taxable"]) if t.get("taxable") is not None else "-",
            ])
        story.append(_section_table(
            ["Enstrüman", "Tarih", "İşlem", "Adet", "Fiyat", "TL Tutar", "Brüt K/Z", "Vergiye Esas"],
            rows,
            [34 * mm, 20 * mm, 14 * mm, 14 * mm, 16 * mm, 26 * mm, 26 * mm, 26 * mm],
        ))
    else:
        story.append(Paragraph("İşlem bulunamadı.", body))

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()
