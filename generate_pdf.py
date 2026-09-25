import os
import io
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfgen import canvas
import tempfile
import re

ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets')
HEADER_IMG = os.path.join(ASSETS_DIR, 'image2.jpeg')
FOOTER_IMG = os.path.join(ASSETS_DIR, 'image3.jpeg')
WATERMARK_IMG = os.path.join(ASSETS_DIR, 'image1.png')

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
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, num_pages):
        page_w, page_h = A4

        # Watermark
        if os.path.exists(WATERMARK_IMG):
            self.saveState()
            self.setFillAlpha(0.08)
            wm_w = 120 * mm
            wm_h = 65 * mm
            self.drawImage(WATERMARK_IMG, (page_w - wm_w) / 2, (page_h - wm_h) / 2, width=wm_w, height=wm_h, mask='auto')
            self.restoreState()

        # Header Image (top banner)
        if os.path.exists(HEADER_IMG):
            hdr_w = page_w - 24 * mm
            hdr_h = 24 * mm
            self.drawImage(HEADER_IMG, 12 * mm, page_h - 30 * mm, width=hdr_w, height=hdr_h, preserveAspectRatio=True)

        # Footer Image (bottom banner)
        if os.path.exists(FOOTER_IMG):
            ftr_w = page_w - 24 * mm
            ftr_h = 18 * mm
            self.drawImage(FOOTER_IMG, 12 * mm, 6 * mm, width=ftr_w, height=ftr_h, preserveAspectRatio=True)

        # Page number indicator removed as requested by user

def build_quote_pdf(quote_data):
    """
    Builds a high quality PDF quotation.
    First tries Word COM to export the exact .docx layout, falling back to ReportLab.
    """
    # Try Word COM for 100% exact rendering of the Word template
    try:
        from generate_word import build_quote_word
        import win32com.client
        import pythoncom

        pythoncom.CoInitialize()
        docx_stream = build_quote_word(quote_data)

        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp_docx:
            tmp_docx.write(docx_stream.read())
            tmp_docx_path = tmp_docx.name

        tmp_pdf_path = tmp_docx_path.replace('.docx', '.pdf')

        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0

        doc = word.Documents.Open(os.path.abspath(tmp_docx_path), ReadOnly=True)
        doc.SaveAs(os.path.abspath(tmp_pdf_path), FileFormat=17) # wdFormatPDF = 17
        doc.Close(False)
        word.Quit()

        with open(tmp_pdf_path, 'rb') as f:
            pdf_bytes = f.read()

        # Clean up temp files
        try: os.remove(tmp_docx_path)
        except: pass
        try: os.remove(tmp_pdf_path)
        except: pass

        buf = io.BytesIO(pdf_bytes)
        buf.seek(0)
        return buf
    except Exception as e:
        print(f"Word COM PDF export fallback triggered: {e}")
        return build_reportlab_pdf(quote_data)

def build_reportlab_pdf(quote_data):
    """
    Fallback ReportLab PDF generator matching demo quotation layout.
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=34 * mm,
        bottomMargin=26 * mm
    )

    styles = getSampleStyleSheet()

    style_client = ParagraphStyle(
        'ClientText',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=12,
        textColor=colors.HexColor('#111111')
    )
    style_ref = ParagraphStyle(
        'RefText',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=12,
        alignment=TA_RIGHT,
        textColor=colors.HexColor('#111111')
    )
    style_subject = ParagraphStyle(
        'SubjectText',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9.5,
        leading=13,
        textColor=colors.HexColor('#111111')
    )
    style_intro = ParagraphStyle(
        'IntroText',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor('#333333')
    )
    style_cell = ParagraphStyle(
        'CellText',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=10
    )
    style_cell_bold = ParagraphStyle(
        'CellBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=10
    )
    style_cell_center = ParagraphStyle(
        'CellCenter',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=10,
        alignment=TA_CENTER
    )
    style_cell_right = ParagraphStyle(
        'CellRight',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=10,
        alignment=TA_RIGHT
    )

    company_name = quote_data.get('company_name', '')
    location = quote_data.get('location', '')
    attn = quote_data.get('attn', '')
    yr_ref = quote_data.get('yr_ref', '')
    quotation_ref = quote_data.get('quotation_ref', 'REE-07-07/23')
    quotation_date = quote_data.get('quotation_date', '')
    subject = quote_data.get('subject', 'Scaffolding Materials on sale basis(Used and Refurbed materials)')
    intro_note = quote_data.get('intro_note', 'We thank you for whatsapp inquiry and pleased quote for used equipment as follows:')
    items = quote_data.get('items', [])
    subtotal = float(quote_data.get('subtotal', 0.0))
    vat_amount = float(quote_data.get('vat_amount', 0.0))
    grand_total = float(quote_data.get('grand_total', 0.0))
    terms = quote_data.get('terms_conditions', '')
    bank_details = quote_data.get('bank_details', '')
    sig_name = quote_data.get('signatory_name', 'Mohamed Faizal')
    sig_title = quote_data.get('signatory_title', 'Rawaiya AL Etihad Est.')

    story = []

    # Client Info + Ref / Date side-by-side Table (Only on Page 1, No Page 1 of 2 indicator)
    client_html = f"<b>M/S. {company_name}</b><br/>{location}<br/>Attn:.: {attn}"
    if yr_ref:
        client_html += f"<br/>Yr. reference : {yr_ref}"
    else:
        client_html += "<br/>Yr. reference :"

    ref_html = f"<b>Date: {quotation_date}</b><br/><b>Our Ref : {quotation_ref}</b>"

    info_table_data = [
        [Paragraph(client_html, style_client), Paragraph(ref_html, style_ref)]
    ]
    info_table = Table(info_table_data, colWidths=[105 * mm, 77 * mm])
    info_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 2 * mm))

    # Subject & Intro
    story.append(Paragraph(f"<b>Sub:</b>   {subject}", style_subject))
    story.append(Spacer(1, 1.5 * mm))
    story.append(Paragraph(intro_note, style_intro))
    story.append(Spacer(1, 2.5 * mm))

    # Items Table
    table_data = [
        [
            Paragraph("<b>SL</b>", style_cell_center),
            Paragraph("<b>Item Description</b>", style_cell_bold),
            Paragraph("<b>Unit</b>", style_cell_center),
            Paragraph("<b>Qty</b>", style_cell_center),
            Paragraph("<b>Unit Rate<br/>(SR)</b>", style_cell_center),
            Paragraph("<b>Amount<br/>(SR)</b>", style_cell_center),
            Paragraph("<b>Remarks</b>", style_cell_center),
        ]
    ]

    for idx, item in enumerate(items, 1):
        qty_val = float(item.get('qty', 0))
        rate_val = float(item.get('rate', 0.0))
        amt_val = float(item.get('amount', qty_val * rate_val))

        table_data.append([
            Paragraph(str(idx), style_cell_center),
            Paragraph(f"<b>{item.get('desc', '')}</b>", style_cell),
            Paragraph(item.get('unit', 'Pcs.'), style_cell_center),
            Paragraph(f"<b>{qty_val:,.0f}</b>" if qty_val.is_integer() else f"<b>{qty_val:,.2f}</b>", style_cell_center),
            Paragraph(f"<b>{rate_val:,.2f}</b>", style_cell_right),
            Paragraph(f"<b>{amt_val:,.2f}</b>", style_cell_right),
            Paragraph(item.get('remarks', ''), style_cell),
        ])

    # Subtotal Row
    table_data.append([
        "", "", "",
        Paragraph("<b>Total   SR</b>", style_cell_bold), "",
        Paragraph(f"<b>{subtotal:,.2f}</b>", style_cell_right),
        ""
    ])
    # VAT Row
    table_data.append([
        "", "", "",
        Paragraph("<b>VAT@15%</b>", style_cell_bold), "",
        Paragraph(f"<b>{vat_amount:,.2f}</b>", style_cell_right),
        ""
    ])
    # Grand Total Row
    table_data.append([
        "", "", "",
        Paragraph("<b>Grand Total</b>", style_cell_bold), "",
        Paragraph(f"<b>{grand_total:,.2f}</b>", style_cell_right),
        ""
    ])

    col_widths = [10 * mm, 62 * mm, 15 * mm, 18 * mm, 24 * mm, 28 * mm, 25 * mm]
    item_table = Table(table_data, colWidths=col_widths, repeatRows=1)

    t_style = [
        ('GRID', (0, 0), (-1, len(items)), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('LEFTPADDING', (0, 0), (-1, -1), 2),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F8F8F8')),
    ]

    num_rows = len(table_data)
    for r_idx in range(num_rows - 3, num_rows):
        t_style.extend([
            ('SPAN', (3, r_idx), (4, r_idx)),
            ('ALIGN', (3, r_idx), (4, r_idx), 'RIGHT'),
            ('GRID', (3, r_idx), (5, r_idx), 0.5, colors.black),
            ('BACKGROUND', (3, r_idx), (5, r_idx), colors.HexColor('#F4F4F4')),
        ])

    item_table.setStyle(TableStyle(t_style))
    story.append(item_table)

    # Page Break for Terms & Conditions (Page 2)
    story.append(PageBreak())

    # Page 2: Terms & Conditions
    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph("<b><u>Terms & Conditions</u></b>", style_subject))
    story.append(Spacer(1, 3 * mm))

    if terms:
        for idx, t_line in enumerate(terms.split('\n'), 1):
            if t_line.strip():
                clean_term = re.sub(r'^\s*(\d+[\.\)]\s*)+', '', t_line.strip())
                story.append(Paragraph(f"<b>{idx}. {clean_term}</b>", style_cell))
                story.append(Spacer(1, 1.5 * mm))

    story.append(Spacer(1, 4 * mm))

    # Bank Details
    if bank_details:
        for b_line in bank_details.split('\n'):
            if b_line.strip():
                clean_bank = re.sub(r'^(Our\s*:\s*)+', '', b_line.strip())
                if 'Bank' in clean_bank and not clean_bank.startswith('Our :'):
                    clean_bank = f"Our : {clean_bank}"
                story.append(Paragraph(f"<b>{clean_bank}</b>", style_cell))
                story.append(Spacer(1, 1.2 * mm))

    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph("Should you require any further information, Please do not hesitate contact us, It will be a pleasure to discuss it", style_intro))
    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph("Thank You and Best regards,", style_cell))
    story.append(Spacer(1, 6 * mm))
    story.append(Paragraph(f"<b>{sig_name}</b>", style_cell_bold))
    story.append(Paragraph(f"<b>{sig_title}</b>", style_cell_bold))

    doc.build(story, canvasmaker=NumberedCanvas)
    buffer.seek(0)
    return buffer
