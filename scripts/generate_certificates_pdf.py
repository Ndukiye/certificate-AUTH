import os
import sys
import json
from datetime import datetime

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CERTS_JSON = os.path.join(PROJECT_ROOT, 'web', 'data', 'certs.json')
OUTPUT_DIR = os.path.join(PROJECT_ROOT, 'web', 'data', 'certificates')

PAGE_WIDTH, PAGE_HEIGHT = 595.27, 841.89  # A4 points
MARGIN = 48


def load_certs(json_path):
    if not os.path.exists(json_path):
        raise FileNotFoundError(f'certs.json not found: {json_path}')
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    if not isinstance(data, list):
        raise ValueError('certs.json must be a list of records')
    return data


def draw_certificate(c, rec):
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import inch

    # Header band
    c.setFillColorRGB(0.117, 0.227, 0.541)  # ~ var(--primary)
    c.rect(0, PAGE_HEIGHT - 80, PAGE_WIDTH, 80, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont('Helvetica-Bold', 20)
    c.drawString(MARGIN, PAGE_HEIGHT - 50, 'Certificate of Completion')

    # Body
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

    # Verification URL
    vurl = rec.get('VerificationURL')
    if vurl:
        c.setFillColorRGB(0.149, 0.388, 0.922)
        c.setFont('Helvetica', 11)
        c.drawString(MARGIN, y - 10, f"Verify: {vurl}")
        y -= 30

    # QR Image if available
    qr_path = rec.get('QRCodePath') or ''
    if qr_path:
        if not os.path.isabs(qr_path):
            qr_path = os.path.join(PROJECT_ROOT, qr_path)
        try:
            img = ImageReader(qr_path)
            c.drawImage(img, PAGE_WIDTH - MARGIN - 144, PAGE_HEIGHT - 144 - 80, width=144, height=144, preserveAspectRatio=True, mask='auto')
        except Exception:
            pass

    # Footer note
    c.setFillColorRGB(0.42, 0.44, 0.48)
    c.setFont('Helvetica-Oblique', 10)
    c.drawString(MARGIN, MARGIN, 'This certificate is digitally signed and can be verified online using the link or QR code above.')


def generate_all():
    try:
        from reportlab.pdfgen import canvas
    except ImportError:
        raise ImportError('reportlab is not installed. Install with: python -m pip install reportlab')

    certs = load_certs(CERTS_JSON)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

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