import os
import sys

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INPUT_RTF = os.path.join(PROJECT_ROOT, 'data', 'Certificates_Merged.rtf')
OUTPUT_PDF = os.path.join(PROJECT_ROOT, 'data', 'Certificates_Merged.pdf')


def main():
    try:
        from win32com.client import gencache
        import pythoncom
        import time
    except ImportError:
        print('pywin32 is not installed. Install with: python -m pip install pywin32', file=sys.stderr)
        sys.exit(2)

    if not os.path.exists(INPUT_RTF):
        print(f'Input RTF not found: {INPUT_RTF}', file=sys.stderr)
        sys.exit(1)

    # Initialize COM and use early-bound Word
    pythoncom.CoInitialize()
    word = gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    # 0 = wdAlertsNone
    word.DisplayAlerts = 0
    time.sleep(0.5)

    doc = None
    try:
        input_abs = os.path.abspath(INPUT_RTF)
        # Retry a few times to avoid callee rejection while Word initializes
        for attempt in range(3):
            try:
                doc = word.Documents.Open(FileName=input_abs, ReadOnly=False, AddToRecentFiles=False, Visible=False)
                break
            except Exception:
                time.sleep(0.5)
        if doc is None:
            raise RuntimeError('Failed to open document for export')

        # Update fields to refresh INCLUDEPICTURE and hyperlinks
        try:
            doc.Fields.Update()
            # Also attempt to update fields across story ranges
            for idx in range(1, 10):
                try:
                    doc.StoryRanges(idx).Fields.Update()
                except Exception:
                    pass
        except Exception:
            try:
                word.ActiveDocument.Fields.Update()
            except Exception:
                pass

        # Export as fixed format PDF (wdExportFormatPDF = 17)
        doc.ExportAsFixedFormat(OutputFileName=OUTPUT_PDF, ExportFormat=17)
        print(f'Exported PDF: {OUTPUT_PDF}')
    except Exception as e:
        print(f'Failed to export PDF: {e}', file=sys.stderr)
        sys.exit(3)
    finally:
        try:
            if doc is not None and hasattr(doc, 'Close'):
                doc.Close(False)
        except Exception:
            pass
        try:
            word.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == '__main__':
    main()