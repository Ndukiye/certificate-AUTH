"""
Certificate Authentication System - Certificate Chaining Script

This script implements a blockchain-like verification system for certificates by:
1. Reading certificate data from a CSV file
2. Creating a chain where each certificate references the hash of the previous one
3. Computing a unique hash for each certificate based on its data and the previous hash
4. Generating verification URLs and QR codes for each certificate
5. Exporting the chained data to CSV, JSON, and creating a hash lookup index

The system ensures certificate authenticity through the unbroken chain of hashes.
"""

import argparse
import csv
import hashlib
import json
import os
from datetime import datetime
import urllib.parse
import urllib.request

# First certificate in chain uses this as its previous hash
GENESIS_HASH = "0" * 64  
# All certificates must contain these fields
REQUIRED_COLUMNS = [
    "CertID",         # Unique identifier for the certificate
    "RecipientName",  # Name of the person receiving the certificate
    "CourseTitle",    # Title of the completed course/program
    "DateIssued",     # Date when the certificate was issued (YYYY-MM-DD)
    "PreviousHash",   # Hash of the previous certificate in the chain
    "CurrentHash",    # Hash of this certificate (computed by this script)
]


def _trim(value):
    """Remove leading/trailing whitespace and handle None values."""
    return (value or "").strip()


def _normalize_date(value):
    """Ensure date is in YYYY-MM-DD format and valid."""
    s = _trim(value)
    try:
        dt = datetime.strptime(s, "%Y-%m-%d")
    except ValueError:
        raise ValueError(f"DateIssued must be ISO format YYYY-MM-DD, got: '{s}'")
    return dt.strftime("%Y-%m-%d")


def _is_hex64(s):
    """Validate that a string is a 64-character hexadecimal value (SHA-256 hash)."""
    s = _trim(s)
    if len(s) != 64:
        return False
    try:
        int(s, 16)
        return True
    except ValueError:
        return False


def _concat_fields(cert_id, name, title, date_issued, previous_hash):
    """Combine certificate fields into a single string for hashing."""
    return f"{cert_id}|{name}|{title}|{date_issued}|{previous_hash}"


def _sha256_hex(s):
    """Compute SHA-256 hash of a string and return as hexadecimal."""
    return hashlib.sha256(s.encode("utf-8")).hexdigest()


def _build_qr_url(data, size=180):
    """Generate a URL to create a QR code using the QR Server API."""
    params = urllib.parse.urlencode({"data": data, "size": f"{size}x{size}"})
    return f"https://api.qrserver.com/v1/create-qr-code/?{params}"


def _download_qr(qr_url, out_path):
    """Download a QR code image from a URL and save it to the specified path."""
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    urllib.request.urlretrieve(qr_url, out_path)


def chain_certificates(input_path, output_path, json_path=None, index_path=None, base_url=None, qr_dir=None, qr_size=180, add_qr_url=False, qr_absolute=False):
    """
    Process certificates and create a verification chain.
    
    Args:
        input_path: Path to input CSV with certificate data
        output_path: Path to save the chained certificate CSV
        json_path: Optional path to save JSON data for web verification
        index_path: Optional path to save hash lookup index
        base_url: Base URL for verification page (adds VerificationURL field)
        qr_dir: Directory to save QR code images
        qr_size: Size of QR codes in pixels
        add_qr_url: Whether to include QR code URLs in output
        qr_absolute: Whether to use absolute paths for QR code files
    """
    # Validate input file exists
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    # Read and validate input CSV
    with open(input_path, "r", newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames or []
        # Check that all required columns are present
        missing = [c for c in REQUIRED_COLUMNS if c not in fieldnames]
        if missing:
            raise ValueError(f"Input CSV missing required columns: {missing}")

        rows = [row for row in reader]

    if not rows:
        raise ValueError("Input CSV has no data rows")

    output_rows = []
    seen_hashes = set()  # Track hashes to prevent duplicates

    # Process each certificate row
    for i, row in enumerate(rows):
        # Extract and clean certificate data
        cert_id = _trim(row.get("CertID"))
        name = _trim(row.get("RecipientName"))
        title = _trim(row.get("CourseTitle"))
        date_issued = _normalize_date(row.get("DateIssued"))

        # Determine previous hash - either from input or from previous certificate
        prev_hash_raw = _trim(row.get("PreviousHash"))
        if i == 0:
            # First certificate uses provided hash or genesis hash
            previous_hash = prev_hash_raw or GENESIS_HASH
        else:
            # Subsequent certificates link to the previous certificate's hash
            previous_hash = _trim(output_rows[i - 1]["CurrentHash"]) or GENESIS_HASH

        # Validate previous hash format
        previous_hash = previous_hash.lower()
        if not _is_hex64(previous_hash):
            raise ValueError(
                f"Row {i+2}: Invalid PreviousHash (must be 64 hex chars): '{previous_hash}'"
            )

        # Create the certificate's unique hash
        concat_str = _concat_fields(cert_id, name, title, date_issued, previous_hash)
        current_hash = _sha256_hex(concat_str)

        # Ensure no duplicate hashes (would indicate duplicate certificates)
        if current_hash in seen_hashes:
            raise ValueError(
                f"Duplicate CurrentHash detected at row {i+2}: {current_hash}. Input data must be unique."
            )
        seen_hashes.add(current_hash)

        # Build the output certificate record
        output_row = {
            "CertID": cert_id,
            "RecipientName": name,
            "CourseTitle": title,
            "DateIssued": date_issued,
            "PreviousHash": previous_hash,
            "CurrentHash": current_hash,
        }

        # Add verification URL if base URL provided
        verification_url = None
        if base_url:
            verification_url = f"{base_url.rstrip('/')}/?hash={current_hash}"
            output_row["VerificationURL"] = verification_url

        # Add QR code URL if requested
        if verification_url and add_qr_url:
            output_row["QRCodeURL"] = _build_qr_url(verification_url, size=qr_size)

        # Generate and download QR code image if directory specified
        if verification_url and qr_dir:
            # Name QR file by certificate ID or sequential number
            qr_filename = f"{cert_id}.png" if cert_id else f"{i+1:03d}.png"
            qr_path = os.path.join(qr_dir, qr_filename)
            try:
                _download_qr(_build_qr_url(verification_url, size=qr_size), qr_path)
                # Store absolute or relative path based on configuration
                output_row["QRCodePath"] = os.path.abspath(qr_path) if qr_absolute else qr_path
            except Exception as e:
                print(f"Warning: Failed to download QR for {cert_id or i+1}: {e}")

        output_rows.append(output_row)

    # Prepare output column list based on enabled features
    out_fieldnames = REQUIRED_COLUMNS[:]
    if base_url:
        out_fieldnames.append("VerificationURL")
    if add_qr_url:
        out_fieldnames.append("QRCodeURL")
    if qr_dir:
        out_fieldnames.append("QRCodePath")

    # Write chained certificates to CSV output
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=out_fieldnames)
        writer.writeheader()
        for row in output_rows:
            writer.writerow(row)

    # Export certificates as JSON for web verification
    if json_path:
        os.makedirs(os.path.dirname(json_path) or ".", exist_ok=True)
        with open(json_path, "w", encoding="utf-8") as jf:
            json.dump(output_rows, jf, ensure_ascii=False, indent=2)

    # Create hash lookup index for fast certificate retrieval
    if index_path:
        os.makedirs(os.path.dirname(index_path) or ".", exist_ok=True)
        index = {row["CurrentHash"]: i for i, row in enumerate(output_rows)}
        with open(index_path, "w", encoding="utf-8") as inf:
            json.dump(index, inf, ensure_ascii=False, indent=2)


def main(base_path=None, input_path=None, output_path=None, json_path=None, index_path=None, base_url=None, qr_dir=None, qr_size=180, add_qr_url=False, qr_absolute=False):
    """
    Main function to chain certificates, callable with explicit arguments or via command-line.
    
    Args:
        base_path: Optional base path to resolve PROJECT_ROOT for script execution context.
        input_path: Path to input CSV with certificate data.
        output_path: Path to save the chained certificate CSV.
        json_path: Optional path to save JSON data for web verification.
        index_path: Optional path to save hash lookup index.
        base_url: Base URL for verification page (adds VerificationURL field).
        qr_dir: Directory to save QR code images.
        qr_size: Size of QR codes in pixels.
        add_qr_url: Whether to include QR code URLs in output.
        qr_absolute: Whether to use absolute paths for QR code files.
    """
    # Determine PROJECT_ROOT based on base_path or current script location
    if base_path:
        PROJECT_ROOT = base_path
    else:
        PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))

    # If arguments are not provided, parse from command line (for script execution)
    if input_path is None or output_path is None:
        parser = argparse.ArgumentParser(
            description="Chain certificates by computing SHA-256 over normalized fields and previous hash"
        )
        # Required arguments
        parser.add_argument("--input", required=True, help="Path to input CSV")
        parser.add_argument(
            "--output", required=True, help="Path to output chained CSV"
        )
        
        # Web verification options
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
        
        # QR code options
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
        input_path = args.input
        output_path = args.output
        json_path = args.json
        index_path = args.index
        base_url = args.base_url
        qr_dir = args.qr_dir
        qr_size = args.qr_size
        add_qr_url = args.add_qr_url
        qr_absolute = args.qr_absolute

    # Define default paths relative to PROJECT_ROOT if not provided
    if input_path is None: input_path = os.path.join(PROJECT_ROOT, 'data', 'Certificates.csv')
    if output_path is None: output_path = os.path.join(PROJECT_ROOT, 'data', 'Certificates_Chained.csv')
    if json_path is None: json_path = os.path.join(PROJECT_ROOT, 'web', 'data', 'certs.json')
    if index_path is None: index_path = os.path.join(PROJECT_ROOT, 'web', 'data', 'hash_index.json')
    if qr_dir is None: qr_dir = os.path.join(PROJECT_ROOT, 'web', 'data', 'qrcodes')

    # Process certificates with parsed arguments
    chain_certificates(
        input_path=input_path,
        output_path=output_path,
        json_path=json_path,
        index_path=index_path,
        base_url=base_url,
        qr_dir=qr_dir,
        qr_size=qr_size,
        add_qr_url=add_qr_url,
        qr_absolute=qr_absolute,
    )


if __name__ == "__main__":
    main()