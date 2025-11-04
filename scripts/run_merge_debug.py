import os
import sys

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'scripts')))
import export_merge_docx

project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
input_csv_path = os.path.join(project_root, 'data', 'Certificates_Chained.csv')
template_path = os.path.join(project_root, 'templates', 'certificate_template.docx')
output_docx_path = os.path.join(project_root, 'output', 'Certificates_Merged.docx')

print(f"Project Root: {project_root}")
print(f"Input CSV Path: {input_csv_path}")
print(f"Template Path: {template_path}")
print(f"Output DOCX Path: {output_docx_path}")

try:
    export_merge_docx.main(
        input_csv_path=input_csv_path,
        template_path=template_path,
        output_docx_path=output_docx_path,
        base_path=project_root
    )
    print("Mail merge executed successfully!")
except Exception as e:
    import traceback
    print(f"Error during mail merge: {e}")
    traceback.print_exc()