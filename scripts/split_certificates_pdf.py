import os
import sys
import json

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MERGED_PDF = os.path.join(PROJECT_ROOT, 'data', 'Certificates_Merged.pdf')
CERTS_JSON = os.path.join(PROJECT_ROOT, 'web', 'data', 'certs.json')
OUTPUT_DIR = os.path.join(PROJECT_ROOT, 'web', 'data', 'certificates')


def load_cert_ids(json_path):
    if not os.path.exists(json_path):
        raise FileNotFoundError(f'certs.json not found: {json_path}')
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    if not isinstance(data, list):
        raise ValueError('certs.json must be a list of certificate records')
    ids = []
    for i, rec in enumerate(data):
        cid = str(rec.get('CertID') or '').strip()
        if not cid:
            # fallback to 3-digit index if missing
            cid = f'{i+1:03d}'
        ids.append(cid)
    return ids


def split_pdf(merged_pdf_path, cert_ids, out_dir):
    try:
        from pypdf import PdfReader, PdfWriter
    except Exception:
        try:
            # Older package name
            from PyPDF2 import PdfReader, PdfWriter
        except Exception:
            raise ImportError('Please install pypdf (preferred) or PyPDF2')

    if not os.path.exists(merged_pdf_path):
        raise FileNotFoundError(f'Merged PDF not found: {merged_pdf_path}')

    reader = PdfReader(merged_pdf_path)
    total_pages = len(reader.pages)
    if total_pages == 0:
        raise ValueError('Merged PDF contains no pages')

    os.makedirs(out_dir, exist_ok=True)

    count = min(total_pages, len(cert_ids))
    if total_pages != len(cert_ids):
        print(f'Warning: pages ({total_pages}) != records ({len(cert_ids)}). Proceeding with {count} items.', file=sys.stderr)

    written = 0
    for i in range(count):
        writer = PdfWriter()
        writer.add_page(reader.pages[i])
        cid = cert_ids[i]
        out_path = os.path.join(out_dir, f'{cid}.pdf')
        with open(out_path, 'wb') as out_f:
            writer.write(out_f)
        written += 1
        print(f'Wrote: {out_path}')
    print(f'Done. Wrote {written} certificate PDFs to: {out_dir}')


def main():
    try:
        cert_ids = load_cert_ids(CERTS_JSON)
        split_pdf(MERGED_PDF, cert_ids, OUTPUT_DIR)
    except Exception as e:
        print(f'Error: {e}', file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()