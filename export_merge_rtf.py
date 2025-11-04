import re
import csv
import os

def rtf_escape(text):
    """Escape special characters for RTF."""
    if not isinstance(text, str):
        text = str(text)
    # Basic RTF escaping
    text = text.replace('\\', '\\\\')
    text = text.replace('{', '\\{')
    text = text.replace('}', '\\}')
    # Handle non-ASCII characters
    def encode_char(match):
        char = match.group(0)
        return f"\\u{ord(char)}?"
    return re.sub(r'[^\x00-\x7f]', encode_char, text)

def apply_template(template_content, data):
    """Apply data to RTF template with simplified field handling."""
    result = template_content
    
    # Replace simple MERGEFIELD
    def replace_mergefield(match):
        field_name = match.group(1).strip()
        if field_name in data:
            return rtf_escape(data[field_name])
        return match.group(0)
    result = re.sub(r'\{\s*MERGEFIELD\s+([^}]+)\s*\}', replace_mergefield, result)
    
    # Replace HYPERLINK fields
    def replace_hyperlink(match):
        field_content = match.group(1)
        # Extract the URL from merged content
        url = field_content.strip()
        if url.startswith('{ MERGEFIELD'):
            field_name = re.search(r'MERGEFIELD\s+([^}]+)', url)
            if field_name and field_name.group(1).strip() in data:
                url = data[field_name.group(1).strip()]
        return f'{{\\field{{\\*\\fldinst{{HYPERLINK "{rtf_escape(url)}"}}}}{{\\fldrslt{{{rtf_escape(url)}}}}}}}'
    result = re.sub(r'\{\s*HYPERLINK\s+"([^}]+)"\s*\}', replace_hyperlink, result)
    
    # Replace INCLUDEPICTURE fields
    def replace_picture(match):
        field_content = match.group(1)
        # Extract the path from merged content
        path = field_content.strip()
        if path.startswith('{ MERGEFIELD'):
            field_name = re.search(r'MERGEFIELD\s+([^}]+)', path)
            if field_name and field_name.group(1).strip() in data:
                path = data[field_name.group(1).strip()]
        return f'{{\\field{{\\*\\fldinst{{INCLUDEPICTURE "{rtf_escape(path)}" \\\\d}}}}{{\\fldrslt{{}}}}}}'
    result = re.sub(r'\{\s*INCLUDEPICTURE\s+"([^}]+)"\s*\}', replace_picture, result)
    
    return result

def main():
    # File paths
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(script_dir, 'templates', 'certificate_template.rtf')
    data_path = os.path.join(script_dir, 'data', 'Certificates_Chained.csv')
    output_path = os.path.join(script_dir, 'output', 'Certificates_Merged.rtf')
    
    # Read template
    with open(template_path, 'r', encoding='utf-8') as f:
        template_content = f.read()
    
    # Read CSV data
    with open(data_path, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        data = next(reader)  # Get first row
    
    # Apply template
    result = apply_template(template_content, data)
    
    # Write output
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(result)

if __name__ == '__main__':
    main()