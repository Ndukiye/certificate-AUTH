Secure Certificate Chain

Overview
- Create cryptographically verifiable certificates by chaining each record’s hash to the previous one (blockchain-like). Any tampering breaks the chain and fails verification.

Structure
- `data/` — input and output datasets
- `scripts/` — Python chaining script
- `web/` — static verification site (HTML/JS/CSS)

Quick Start
1) Prepare input CSV: `data/Certificates.csv`
   - Columns: `CertID, RecipientName, CourseTitle, DateIssued, PreviousHash, CurrentHash`
   - For the first row (genesis), set `PreviousHash` to 64 zeros.
   - Use ISO date format: `YYYY-MM-DD`.

2) Generate chained outputs
   - Run from project root:
     - `python scripts/chain_certificates.py --input data/Certificates.csv --output data/Certificates_Chained.csv --json web/data/certs.json --index web/data/hash_index.json --base-url http://localhost:8000/`

   Optional: Include QR support (Phase 3A)
   - Add an online QR image URL per certificate:
     - `python scripts/chain_certificates.py --input data/Certificates.csv --output data/Certificates_Chained.csv --json web/data/certs.json --index web/data/hash_index.json --base-url http://localhost:8000/ --add-qr-url --qr-size 180`
   - Download QR PNGs locally (saved to `data/qrcodes/CertID.png`):
     - `python scripts/chain_certificates.py --input data/Certificates.csv --output data/Certificates_Chained.csv --json web/data/certs.json --index web/data/hash_index.json --base-url http://localhost:8000/ --qr-dir data/qrcodes --qr-size 180`
   - Notes:
     - QR images are generated via `https://api.qrserver.com/v1/create-qr-code/` using the `VerificationURL` as data.
     - `--add-qr-url` adds `QRCodeURL` to CSV/JSON; `--qr-dir` saves PNGs and adds `QRCodePath`.

3) Serve verification site
   - Change directory to `web/` and run:
     - `python -m http.server 8000`
   - Open `http://localhost:8000/` and paste a `CurrentHash` to verify.

Notes
- Deterministic hashing: concatenation order is `CertID|RecipientName|CourseTitle|DateIssued|PreviousHash` and SHA-256 is computed over UTF‑8 bytes.
- The site recalculates the hash client-side and compares to stored `CurrentHash`.
- UI validation: input must be 64 hex characters; chain link checks warn if the previous record is missing or mismatched.


## Mail Merge (Word) with QR Codes
- Guide: see `docs/mail_merge_guide.md` for step-by-step instructions.
- Recommended flags when generating data:
  - `--qr-dir data/qrcodes` to save per-record QR PNGs
  - `--qr-size 180` to size the QR image
  - `--add-qr-url` to include `QRCodeURL` in outputs
  - `--qr-absolute` so `QRCodePath` is an absolute path (Word resolves these reliably)
- In Word, embed QR images with:
  - Insert field code via `Ctrl+F9`: `INCLUDEPICTURE "{ MERGEFIELD QRCodePath }" \d`
  - After finishing merge to a new document, press `Ctrl+A` then `F9` to refresh images.


## Production Hosting (Vercel)
- Requirements: Node.js and Vercel CLI (`npm i -g vercel`).
- Steps:
  - Run `vercel login` to authenticate.
  - From repo root, run `vercel --prod --confirm`.
  - Vercel will deploy and return your URL (e.g., `https://certificate-auth.vercel.app`).
- Regenerate data with the Vercel URL so QR links and verification pages point correctly:
  - Example:
    - `python scripts/chain_certificates.py --input data/Certificates.csv --output data/Certificates_Chained.csv --json web/data/certs.json --index web/data/hash_index.json --base-url https://<your-project>.vercel.app/ --qr-dir data/qrcodes --qr-size 180 --add-qr-url --qr-absolute`
  - Then rebuild the merged RTF:
    - `python scripts/export_merge_rtf.py`
- Notes:
  - `vercel.json` maps site root to `web/` so assets resolve at `/`.
  - For a custom domain, add it in the Vercel dashboard and re-run the chain script with that domain as `--base-url`.

## Production Hosting (GitHub Pages)
- This repo includes `.github/workflows/deploy.yml` to auto-deploy the `web/` folder to GitHub Pages.
- Steps:
  - Create a GitHub repository and push this project.
  - Ensure your default branch is `main` (or adjust workflow branches accordingly).
  - In GitHub: Settings → Pages → Source: `GitHub Actions` (the workflow will publish on push).
- Base URL for QR `VerificationURL`:
  - User/Org Pages: `https://<username>.github.io/`
  - Project Pages: `https://<username>.github.io/<repo>/`
- Regenerate data with the production base URL so QR links point to your hosted site:
  - Example (Project Pages):
    - `python scripts/chain_certificates.py --input data/Certificates.csv --output data/Certificates_Chained.csv --json web/data/certs.json --index web/data/hash_index.json --base-url https://<username>.github.io/<repo>/ --qr-dir data/qrcodes --qr-size 180 --add-qr-url --qr-absolute`
- After pushing to `main`, the action publishes `web/` and you can verify at your Pages URL.


### Ready-made Mail Merge Template
- Template: `templates/certificate_template.rtf` (opens in Word; you can Save As `.docx`).
- Usage:
  - Open the template in Word.
  - Go to `Mailings` > `Select Recipients` > `Use an Existing List...` and choose `data/Certificates_Chained.csv`.
  - Preview results; the QR image uses `QRCodePath` and the link uses `VerificationURL`.
  - Finish & Merge to a new document, then press `Ctrl+A` and `F9` to refresh QR images.
- Tip: Generate data with `--qr-absolute` so `QRCodePath` resolves reliably.