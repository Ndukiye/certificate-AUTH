import zipfile
import re

def main(path='output/Certificates_Merged.docx'):
    with zipfile.ZipFile(path) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    has_nested = bool(re.search(r'MERGEFIELD\s*QRCodePath', xml))
    m = re.search(r'INCLUDEPICTURE[\s\S]{0,300}', xml)
    print('Nested MERGEFIELD remaining:', has_nested)
    print('First INCLUDEPICTURE snippet:\n', (m.group(0)[:300] if m else 'None'))

if __name__ == '__main__':
    main()