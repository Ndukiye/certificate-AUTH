import csv
from mailmerge import MailMerge
import os
import sys
import argparse


def main(input_csv_path=None, template_path=None, output_docx_path=None, base_path=None):
    print(f"export_merge_docx.py main called with: input_csv_path={input_csv_path}, template_path={template_path}, output_docx_path={output_docx_path}, base_path={base_path}")

    if base_path is None:
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

    if input_csv_path is None:
        input_csv_path = os.path.join(base_path, 'data', 'Certificates_Chained.csv')
    if template_path is None:
        template_path = os.path.join(base_path, 'templates', 'certificate_template.docx')
    if output_docx_path is None:
        output_docx_path = os.path.join(base_path, 'data', 'Certificates_Merged.docx')

    # Ensure paths are absolute and correctly resolved
    input_csv_path = os.path.abspath(input_csv_path)
    template_path = os.path.abspath(template_path)
    output_docx_path = os.path.abspath(output_docx_path)
    base_path = os.path.abspath(base_path)

    print(f"Resolved paths: input_csv_path={input_csv_path}, template_path={template_path}, output_docx_path={output_docx_path}, base_path={base_path}")

    if not os.path.exists(input_csv_path):
        msg = f"Input CSV file not found at {input_csv_path}"
        print(f"Error: {msg}")
        raise FileNotFoundError(msg)
    if not os.path.exists(template_path):
        msg = f"Template DOCX file not found at {template_path}"
        print(f"Error: {msg}")
        raise FileNotFoundError(msg)

    try:
        with open(input_csv_path, 'r', encoding='utf-8') as csv_file:
            csv_reader = csv.DictReader(csv_file)
            records = list(csv_reader)

        if not records:
            print("No records found in the CSV file. Writing original template to output and exiting.")
            # Write the original template to the output path to avoid frontend download issues
            document = MailMerge(template_path)
            # Ensure the output directory exists
            output_dir = os.path.dirname(output_docx_path)
            os.makedirs(output_dir, exist_ok=True)
            document.write(output_docx_path)
            print(f"No data to merge. Output saved to {output_docx_path}")
            return

        # Assuming all records use the same template
        document = MailMerge(template_path)
        
        # Get all merge fields in the document
        merge_fields = document.get_merge_fields()
        print(f"Merge fields found in template: {merge_fields}")

        # Debugging: Print the raw merge fields from the document
        print(f"Raw merge fields from document: {document.get_merge_fields()}")

        # Prepare a list of dictionaries for merging
        merge_data = []
        for record in records:
            # Create a dictionary for each record, ensuring all merge fields are present
            record_data = {}
            for field in merge_fields:
                # Map template merge fields to CSV column headers
                csv_column_map = {
                    'ID': 'CertID',
                    'CompletionDate': 'DateIssued',
                    'Hash': 'CurrentHash',
                    'Link': 'VerificationURL',
                    'QR': 'QRCodeURL',
                    'RecipientName': 'RecipientName',
                    'CourseTitle': 'CourseTitle'
                }
                csv_column_name = csv_column_map.get(field, field) # Use mapping, or field name if not mapped
                
                if field == 'QR':
                    qr_relative_path = record.get('QRCodePath', '')
                    if qr_relative_path:
                        record_data[field] = os.path.join(base_path, qr_relative_path)
                    else:
                        record_data[field] = ''
                else:
                    record_data[field] = record.get(csv_column_name, '') # Use empty string if column not in CSV
            
            # Explicitly add QR field if not already present (e.g., if get_merge_fields() doesn't find it)
            if 'QR' not in merge_fields:
                qr_relative_path = record.get('QRCodePath', '')
                if qr_relative_path:
                    record_data['QR'] = os.path.join(base_path, qr_relative_path)
                else:
                    record_data['QR'] = ''
            merge_data.append(record_data)

        # Merge all records
        document.merge_templates(merge_data, separator='page_break')

        # Ensure the output directory exists
        output_dir = os.path.dirname(output_docx_path)
        os.makedirs(output_dir, exist_ok=True)

        document.write(output_docx_path)
        print(f"Mail merge completed successfully. Output saved to {output_docx_path}")

    except Exception as e:
        import traceback
        print("An error occurred during mail merge:\n" + traceback.format_exc())
        # Re-raise so the Flask route can handle and return a proper 500 JSON
        raise


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Perform mail merge from CSV to DOCX template.')
    parser.add_argument('--input-csv', type=str, help='Path to the input CSV file.')
    parser.add_argument('--template', type=str, help='Path to the DOCX template file.')
    parser.add_argument('--output-docx', type=str, help='Path to the output DOCX file.')
    parser.add_argument('--base-path', type=str, help='Base path for resolving relative paths, e.g., for QR codes.')
    args = parser.parse_args()

    main(args.input_csv, args.template, args.output_docx, args.base_path)