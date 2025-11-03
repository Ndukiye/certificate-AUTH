"""
Certificate PDF Generator

This script generates individual PDF certificates from certificate data in certs.json.
Each certificate includes recipient information, course details, and a QR code for verification.

Usage:
    python generate_certificates_pdf.py

Requirements:
    - ReportLab library (pip install reportlab)
    - Certificate data in web/data/certs.json
"""

import os
import sys
import json
from datetime import datetime

# Path configuration
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CERTS_JSON = os.path.join(PROJECT_ROOT, 'web', 'data', 'certs.json')
OUTPUT_DIR = os.path.join(PROJECT_ROOT, 'web', 'data', 'certificates')

# PDF layout constants
PAGE_WIDTH, PAGE_HEIGHT = 595.27, 841.89  # A4 points
MARGIN = 48


def load_certs(json_path):
    """
    Load certificate data from JSON file
    
    Args:
        json_path (str): Path to the certificates JSON file
        
    Returns:
        list: List of certificate records
        
    Raises:
        FileNotFoundError: If certs.json doesn't exist
        ValueError: If certs.json isn't a list
    """
    if not os.path.exists(json_path):
        raise FileNotFoundError(f'certs.json not found: {json_path}')
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    if not isinstance(data, list):
        raise ValueError('certs.json must be a list of records')
    return data


def draw_certificate(c, rec):
    """
    Draw certificate content on the PDF canvas
    
    Args:
        c (Canvas): ReportLab canvas object to draw on
        rec (dict): Certificate record with recipient and course information
    """
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import inch

    # Header band - blue title bar
    c.setFillColorRGB(0.117, 0.227, 0.541)  # ~ var(--primary)
    c.rect(0, PAGE_HEIGHT - 80, PAGE_WIDTH, 80, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont('Helvetica-Bold', 20)
    c.drawString(MARGIN, PAGE_HEIGHT - 50, 'Certificate of Completion')

    # Body - certificate information
    c.setFillColorRGB(0.12, 0.14, 0.2)
    c.setFont('Helvetica', 12)
    y = PAGE_HEIGHT - 120
    lines = [
        f"Recipient: {rec.get('RecipientName','')}",
        f"Course: {rec.get('CourseTitle','')}",
        f"Date Issued: {rec.get('DateIssued','')}",
        f"CertID: {rec.get('CertID','')}",
    ]
    for line in lines:
        c.drawString(MARGIN, y, line)
        y -= 20

    # Verification URL - link to online verification
    vurl = rec.get('VerificationURL')
    if vurl:
        c.setFillColorRGB(0.149, 0.388, 0.922)
        c.setFont('Helvetica', 11)
        c.drawString(MARGIN, y - 10, f"Verify: {vurl}")
        y -= 30

    # QR Image - add QR code for verification if available
    qr_path = rec.get('QRCodePath') or ''
    if qr_path:
        if not os.path.isabs(qr_path):
            qr_path = os.path.join(PROJECT_ROOT, qr_path)
        try:
            img = ImageReader(qr_path)
            c.drawImage(img, PAGE_WIDTH - MARGIN - 144, PAGE_HEIGHT - 144 - 80, width=144, height=144, preserveAspectRatio=True, mask='auto')
        except Exception:
            pass

    # Footer note - verification instructions
    c.setFillColorRGB(0.42, 0.44, 0.48)
    c.setFont('Helvetica-Oblique', 10)
    c.drawString(MARGIN, MARGIN, 'This certificate is digitally signed and can be verified online using the link or QR code above.')


def generate_all():
    """
    Main function to generate all certificate PDFs from certs.json
    
    Loads certificate data, creates output directory if needed,
    and generates individual PDF files for each certificate.
    """
    try:
        from reportlab.pdfgen import canvas
    except ImportError:
        raise ImportError('reportlab is not installed. Install with: python -m pip install reportlab')

    # Load certificate data and ensure output directory exists
    certs = load_certs(CERTS_JSON)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Generate a PDF for each certificate
    count = 0
    for rec in certs:
        cid = str(rec.get('CertID') or '').strip() or 'certificate'
        out_path = os.path.join(OUTPUT_DIR, f'{cid}.pdf')
        c = canvas.Canvas(out_path, pagesize=(PAGE_WIDTH, PAGE_HEIGHT))
        draw_certificate(c, rec)
        c.showPage()
        c.save()
        print(f'Wrote: {out_path}')
        count += 1
    print(f'Done. Wrote {count} certificate PDFs to: {OUTPUT_DIR}')


if __name__ == '__main__':
    try:
        generate_all()
    except Exception as e:
        print(f'Error: {e}', file=sys.stderr)
        sys.exit(1)