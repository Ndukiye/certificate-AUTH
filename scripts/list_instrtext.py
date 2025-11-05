import zipfile
import re
import argparse

def main(path='output/Certificates_Merged.docx', context_lines=10):
    with zipfile.ZipFile(path) as z:
        xml = z.read('word/document.xml').decode('utf-8')

    # Find all occurrences of INCLUDEPICTURE or potential image embedding XML
    # This pattern is a placeholder and might need refinement based on actual XML structure
    # We're looking for w:instrText that contains INCLUDEPICTURE, or w:drawing elements
    # that indicate an embedded image.
    # For now, let's focus on finding the old INCLUDEPICTURE fields to see if they are gone.
    pattern = re.compile(r'(<w:r>.*?<w:instrText[^>]*>(?:(?!</w:r>).)*INCLUDEPICTURE(?:(?!</w:r>).)*</w:instrText>.*?</w:r>|<w:drawing>.*?</w:drawing>)', re.DOTALL)

    matches = list(pattern.finditer(xml))

    if not matches:
        print("No INCLUDEPICTURE fields or w:drawing elements found in document.xml.")
        # If no matches, let's try to find the text 'MERGEFIELD QRCodePath' to see if it's still there
        if 'MERGEFIELD QRCodePath' in xml:
            print("'MERGEFIELD QRCodePath' found in document.xml, but not within expected INCLUDEPICTURE or drawing tags.")
        return

    print(f"Found {len(matches)} potential image-related XML blocks.")
    for i, m in enumerate(matches):
        print(f"\n--- Match {i+1} ---")
        # Extract a larger context around the match
        start = max(0, m.start() - 500) # 500 characters before
        end = min(len(xml), m.end() + 500) # 500 characters after
        context = xml[start:end]
        print(context)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='List instrText and drawing elements from a DOCX document.')
    parser.add_argument('--path', type=str, default='output/Certificates_Merged.docx', help='Path to the DOCX file.')
    parser.add_argument('--context-lines', type=int, default=10, help='Number of lines of context to show around matches.')
    args = parser.parse_args()
    main(args.path, args.context_lines)