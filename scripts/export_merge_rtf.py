import csv
import os
import sys

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INPUT_CSV = os.path.join(PROJECT_ROOT, 'data', 'Certificates_Chained.csv')
TEMPLATE_RTF = os.path.join(PROJECT_ROOT, 'templates', 'certificate_template.rtf')
OUTPUT_RTF = os.path.join(PROJECT_ROOT, 'data', 'Certificates_Merged.rtf')


def rtf_escape(text: str) -> str:
    if text is None:
        return ''
    return str(text).replace('\\', '\\\\').replace('{', '\\{').replace('}', '\\}')


def apply_template(template_text: str, rec: dict) -> str:
    # Values
    name = rtf_escape(rec.get('RecipientName', ''))
    title = rtf_escape(rec.get('CourseTitle', ''))
    date_issued = rtf_escape(rec.get('DateIssued', ''))
    cert_id = rtf_escape(rec.get('CertID', ''))
    current_hash = rtf_escape(rec.get('CurrentHash', ''))
    verification_url = rec.get('VerificationURL', '') or ''
    qr_path = rec.get('QRCodePath', '') or ''

    out = template_text
    # Replace fldrslt placeholders for MERGEFIELD values
    out = out.replace('\\fldrslt RecipientName', '\\fldrslt ' + name)
    out = out.replace('\\fldrslt CourseTitle', '\\fldrslt ' + title)
    out = out.replace('\\fldrslt DateIssued', '\\fldrslt ' + date_issued)
    out = out.replace('\\fldrslt CertID', '\\fldrslt ' + cert_id)
    out = out.replace('\\fldrslt CurrentHash', '\\fldrslt ' + current_hash)
    out = out.replace('\\fldrslt VerificationURL', '\\fldrslt ' + rtf_escape(verification_url))

    # Replace HYPERLINK target inside field instruction
    if verification_url:
        out = out.replace('HYPERLINK "{ MERGEFIELD VerificationURL }"', 'HYPERLINK "' + rtf_escape(verification_url) + '"')
    # Replace INCLUDEPICTURE path inside field instruction
    if qr_path:
        out = out.replace('INCLUDEPICTURE "{ MERGEFIELD QRCodePath }" \\d', 'INCLUDEPICTURE "' + rtf_escape(qr_path) + '" \\d')

    return out


def main():
    if not os.path.exists(INPUT_CSV):
        print(f'Input CSV not found: {INPUT_CSV}', file=sys.stderr)
        sys.exit(1)
    if not os.path.exists(TEMPLATE_RTF):
        print(f'Template RTF not found: {TEMPLATE_RTF}', file=sys.stderr)
        sys.exit(2)

    with open(TEMPLATE_RTF, 'r', encoding='utf-8') as tf:
        template_text = tf.read()

    with open(INPUT_CSV, 'r', encoding='utf-8', newline='') as f:
        reader = csv.DictReader(f)
        rows = list(reader)

    merged_parts = []
    for i, rec in enumerate(rows):
        page_rtf = apply_template(template_text, rec)
        merged_parts.append(page_rtf)
        if i < len(rows) - 1:
            merged_parts.append('\\page\n')

    os.makedirs(os.path.dirname(OUTPUT_RTF), exist_ok=True)
    with open(OUTPUT_RTF, 'w', encoding='utf-8') as out:
        out.write(''.join(merged_parts))

    print(f'Wrote merged RTF using template: {OUTPUT_RTF}')


if __name__ == '__main__':
    main()