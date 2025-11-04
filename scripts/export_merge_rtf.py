import csv
from mailmerge import MailMerge
import os
import sys

def main(input_csv_path=None, template_path=None, output_rtf_path=None):
    print(f"export_merge_rtf.py main called with: input_csv_path={input_csv_path}, template_path={template_path}, output_rtf_path={output_rtf_path}")

    if input_csv_path is None:
        input_csv_path = os.path.join(os.path.dirname(__file__), '..', 'data', 'Certificates_Chained.csv')
    if template_path is None:
        template_path = os.path.join(os.path.dirname(__file__), '..', 'templates', 'certificate_template.rtf')
    if output_rtf_path is None:
        output_rtf_path = os.path.join(os.path.dirname(__file__), '..', 'data', 'Certificates_Merged.rtf')

    # Ensure paths are absolute and correctly resolved
    input_csv_path = os.path.abspath(input_csv_path)
    template_path = os.path.abspath(template_path)
    output_rtf_path = os.path.abspath(output_rtf_path)

    print(f"Resolved paths: input_csv_path={input_csv_path}, template_path={template_path}, output_rtf_path={output_rtf_path}")

    if not os.path.exists(input_csv_path):
        print(f"Error: Input CSV file not found at {input_csv_path}")
        sys.exit(1)
    if not os.path.exists(template_path):
        print(f"Error: Template RTF file not found at {template_path}")
        sys.exit(1)

    try:
        with open(input_csv_path, 'r', encoding='utf-8') as csv_file:
            csv_reader = csv.DictReader(csv_file)
            records = list(csv_reader)

        if not records:
            print("No records found in the CSV file. Skipping mail merge.")
            # Create an empty RTF file to indicate no merge was performed
            with open(output_rtf_path, 'w', encoding='utf-8') as f:
                f.write("{\\rtf1\\ansi\\deff0 No certificates merged.}")
            return

        # Assuming all records use the same template
        document = MailMerge(template_path)
        
        # Get all merge fields in the document
        merge_fields = document.get_merge_fields()
        print(f"Merge fields found in template: {merge_fields}")

        # Prepare a list of dictionaries for merging
        merge_data = []
        for record in records:
            # Create a dictionary for each record, ensuring all merge fields are present
            record_data = {}
            for field in merge_fields:
                record_data[field] = record.get(field, '') # Use empty string if field not in CSV
            merge_data.append(record_data)

        # Merge all records
        document.merge_templates(merge_data, separator='page_break')

        # Ensure the output directory exists
        output_dir = os.path.dirname(output_rtf_path)
        os.makedirs(output_dir, exist_ok=True)

        document.write(output_rtf_path)
        print(f"Mail merge completed successfully. Output saved to {output_rtf_path}")

    except Exception as e:
        print(f"An error occurred during mail merge: {e}")
        sys.exit(1)

if __name__ == '__main__':
    # This part is for direct script execution, not used by the Flask app
    # You can add argument parsing here if you want to test it standalone
    main()