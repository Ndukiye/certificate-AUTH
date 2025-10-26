import argparse
import csv
import hashlib
import json
import os
from datetime import datetime
import urllib.parse
import urllib.request

GENESIS_HASH = "0" * 64
REQUIRED_COLUMNS = [
    "CertID",
    "RecipientName",
    "CourseTitle",
    "DateIssued",
    "PreviousHash",
    "CurrentHash",
]


def _trim(value):
    return (value or "").strip()


def _normalize_date(value):
    s = _trim(value)
    try:
        dt = datetime.strptime(s, "%Y-%m-%d")
    except ValueError:
        raise ValueError(f"DateIssued must be ISO format YYYY-MM-DD, got: '{s}'")
    return dt.strftime("%Y-%m-%d")


def _is_hex64(s):
    s = _trim(s)
    if len(s) != 64:
        return False
    try:
        int(s, 16)
        return True
    except ValueError:
        return False


def _concat_fields(cert_id, name, title, date_issued, previous_hash):
    return f"{cert_id}|{name}|{title}|{date_issued}|{previous_hash}"


def _sha256_hex(s):
    return hashlib.sha256(s.encode("utf-8")).hexdigest()


def _build_qr_url(data, size=180):
    params = urllib.parse.urlencode({"data": data, "size": f"{size}x{size}"})
    return f"https://api.qrserver.com/v1/create-qr-code/?{params}"


def _download_qr(qr_url, out_path):
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    urllib.request.urlretrieve(qr_url, out_path)


def chain_certificates(input_path, output_path, json_path=None, index_path=None, base_url=None, qr_dir=None, qr_size=180, add_qr_url=False, qr_absolute=False):
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    with open(input_path, "r", newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames or []
        missing = [c for c in REQUIRED_COLUMNS if c not in fieldnames]
        if missing:
            raise ValueError(f"Input CSV missing required columns: {missing}")

        rows = [row for row in reader]

    if not rows:
        raise ValueError("Input CSV has no data rows")

    output_rows = []
    seen_hashes = set()

    for i, row in enumerate(rows):
        cert_id = _trim(row.get("CertID"))
        name = _trim(row.get("RecipientName"))
        title = _trim(row.get("CourseTitle"))
        date_issued = _normalize_date(row.get("DateIssued"))

        prev_hash_raw = _trim(row.get("PreviousHash"))
        if i == 0:
            previous_hash = prev_hash_raw or GENESIS_HASH
        else:
            previous_hash = _trim(output_rows[i - 1]["CurrentHash"]) or GENESIS_HASH

        previous_hash = previous_hash.lower()
        if not _is_hex64(previous_hash):
            raise ValueError(
                f"Row {i+2}: Invalid PreviousHash (must be 64 hex chars): '{previous_hash}'"
            )

        concat_str = _concat_fields(cert_id, name, title, date_issued, previous_hash)
        current_hash = _sha256_hex(concat_str)

        if current_hash in seen_hashes:
            raise ValueError(
                f"Duplicate CurrentHash detected at row {i+2}: {current_hash}. Input data must be unique."
            )
        seen_hashes.add(current_hash)

        output_row = {
            "CertID": cert_id,
            "RecipientName": name,
            "CourseTitle": title,
            "DateIssued": date_issued,
            "PreviousHash": previous_hash,
            "CurrentHash": current_hash,
        }

        verification_url = None
        if base_url:
            verification_url = f"{base_url.rstrip('/')}/?hash={current_hash}"
            output_row["VerificationURL"] = verification_url

        # Optional QR additions
        if verification_url and add_qr_url:
            output_row["QRCodeURL"] = _build_qr_url(verification_url, size=qr_size)

        if verification_url and qr_dir:
            qr_filename = f"{cert_id}.png" if cert_id else f"{i+1:03d}.png"
            qr_path = os.path.join(qr_dir, qr_filename)
            try:
                _download_qr(_build_qr_url(verification_url, size=qr_size), qr_path)
                output_row["QRCodePath"] = os.path.abspath(qr_path) if qr_absolute else qr_path
            except Exception as e:
                print(f"Warning: Failed to download QR for {cert_id or i+1}: {e}")

        output_rows.append(output_row)

    out_fieldnames = REQUIRED_COLUMNS[:]
    if base_url:
        out_fieldnames.append("VerificationURL")
    if add_qr_url:
        out_fieldnames.append("QRCodeURL")
    if qr_dir:
        out_fieldnames.append("QRCodePath")

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=out_fieldnames)
        writer.writeheader()
        for row in output_rows:
            writer.writerow(row)

    if json_path:
        os.makedirs(os.path.dirname(json_path) or ".", exist_ok=True)
        with open(json_path, "w", encoding="utf-8") as jf:
            json.dump(output_rows, jf, ensure_ascii=False, indent=2)

    if index_path:
        os.makedirs(os.path.dirname(index_path) or ".", exist_ok=True)
        index = {row["CurrentHash"]: i for i, row in enumerate(output_rows)}
        with open(index_path, "w", encoding="utf-8") as inf:
            json.dump(index, inf, ensure_ascii=False, indent=2)


def main():
    parser = argparse.ArgumentParser(
        description="Chain certificates by computing SHA-256 over normalized fields and previous hash"
    )
    parser.add_argument("--input", required=True, help="Path to input CSV")
    parser.add_argument(
        "--output", required=True, help="Path to output chained CSV"
    )
    parser.add_argument(
        "--json", required=False, help="Path to JSON export of records (for web verification)"
    )
    parser.add_argument(
        "--index", required=False, help="Path to JSON hash index (hash -> row index)"
    )
    parser.add_argument(
        "--base-url",
        required=False,
        help="Base URL for verification page; adds VerificationURL column as <base-url>/?hash=<CurrentHash>",
    )
    parser.add_argument(
        "--qr-dir",
        required=False,
        help="Directory to save QR PNGs per certificate (requires --base-url)",
    )
    parser.add_argument(
        "--qr-size",
        required=False,
        type=int,
        default=180,
        help="Square size of QR code image in pixels (default: 180)",
    )
    parser.add_argument(
        "--add-qr-url",
        action="store_true",
        help="Include QRCodeURL column pointing to an online QR image (requires --base-url)",
    )
    parser.add_argument(
        "--qr-absolute",
        action="store_true",
        help="Store QRCodePath as an absolute path (recommended for Word Mail Merge)",
    )

    args = parser.parse_args()

    chain_certificates(
        input_path=args.input,
        output_path=args.output,
        json_path=args.json,
        index_path=args.index,
        base_url=args.base_url,
        qr_dir=args.qr_dir,
        qr_size=args.qr_size,
        add_qr_url=args.add_qr_url,
        qr_absolute=args.qr_absolute,
    )


if __name__ == "__main__":
    main()