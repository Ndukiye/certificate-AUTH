# Mail Merge Guide (Word) for Certificate QR Codes

This guide shows how to use `data/Certificates_Chained.csv` to produce printable certificates in Microsoft Word with embedded QR images and verification links.

## Prerequisites
- Microsoft Word (2016+ recommended)
- Generated data via the chaining script:
  - CSV: `data/Certificates_Chained.csv`
  - QR PNGs: `data/qrcodes/*.png` (created by the script)
- Recommended: run the script with `--qr-absolute` so `QRCodePath` contains absolute paths, which Word resolves more reliably.

Example command:
```
python scripts/chain_certificates.py \
  --input data/Certificates.csv \
  --output data/Certificates_Chained.csv \
  --json web/data/certs.json \
  --index web/data/hash_index.json \
  --base-url http://localhost:8000/ \
  --qr-dir data/qrcodes \
  --qr-size 180 \
  --add-qr-url \
  --qr-absolute
```

## Link the CSV as a Data Source
1. Open a new Word document (your certificate layout).
2. Go to `Mailings` > `Start Mail Merge` > choose `Letters`.
3. Click `Select Recipients` > `Use an Existing List...` and choose `data/Certificates_Chained.csv`.
4. Confirm the delimiter and that Word detects the header row.

## Insert Merge Fields
- Place text fields in the certificate body using `Mailings` > `Insert Merge Field`:
  - `RecipientName` (or your name field)
  - `CertID`
  - `CourseTitle`
  - `DateIssued`
  - `CurrentHash` (optional)
  - `VerificationURL` (see hyperlink step below)

## Insert QR Image (Recommended: QRCodePath)
Word can embed images per-record using the `INCLUDEPICTURE` field code.

1. Position the cursor where the QR image should appear.
2. Press `Ctrl+F9` to insert field code braces `{ }` (do not type braces manually).
3. Between the braces, type:
   ```
   INCLUDEPICTURE "{ MERGEFIELD QRCodePath }" \d
   ```
4. Press `Alt+F9` to toggle field code view off.
5. Use `Mailings` > `Preview Results` to preview.
6. When finishing, choose `Finish & Merge` > `Edit Individual Documents`. In the new document, press `Ctrl+A` then `F9` to force images to update for all records.

Notes:
- `\\d` tells Word to delay loading until fields are updated.
- If images do not appear, ensure `QRCodePath` points to a valid PNG. Absolute paths (`--qr-absolute`) are much more reliable.
- Keep the Word template near the project root if using relative paths, so Word can resolve them.

## Insert Verification Hyperlink (VerificationURL)
- Select the text that should act as a link, e.g., `Verify Online`.
- Press `Ctrl+K` (Insert Hyperlink).
- In `Address`, click `Mailings` > `Insert Merge Field` and choose `VerificationURL`.
- Confirm and preview. Each record will link to its unique verification page.

## Optional: Use QRCodeURL Instead of Local PNGs
- You can embed online QR images by using the `INCLUDEPICTURE` field with `QRCodeURL`:
  ```
  INCLUDEPICTURE "{ MERGEFIELD QRCodeURL }" \d
  ```
- This depends on Word’s ability to fetch remote images during merge, which can be less reliable and slower. Prefer local `QRCodePath`.

## Layout Tips
- Fix the QR image size by setting the container to a fixed width/height. Word will scale the image after `F9` updates.
- Use tables or positioned text boxes to keep fields aligned.
- Place `CurrentHash` and `CertID` in small font near the bottom for auditability.

## Troubleshooting
- No image after merge: toggle field codes (`Alt+F9`), select all (`Ctrl+A`), update fields (`F9`).
- Paths not resolving: re-run the script with `--qr-absolute` or move the template to the project root.
- Wrong rows: confirm the CSV is the latest and that you selected the correct data source.
- Preview looks fine but final missing images: always `Finish & Merge` to a new document and do `Ctrl+A`, `F9` there.

## Field Reference (from CSV)
- `RecipientName`, `CertID`, `CourseTitle`, `DateIssued`
- `CurrentHash`, `PreviousHash`
- `VerificationURL`, `QRCodeURL`, `QRCodePath`

You’re now set to produce QR-enabled certificates using Mail Merge. If you want a prebuilt `.docx` template, I can generate one tailored to your CSV headers.