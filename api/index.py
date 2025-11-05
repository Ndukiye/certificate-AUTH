from flask import Flask, render_template, send_from_directory, jsonify, request, session, redirect, url_for
import os
import subprocess
import sys
import csv
import uuid
import functools

# Add the scripts directory to the Python path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'scripts')))

# Import your scripts
import chain_certificates
import generate_certificates_pdf
import export_merge_docx

app = Flask(__name__, static_folder='../web', static_url_path='/', template_folder='../web')
app.secret_key = 'your_very_secret_key_here' # Replace with a strong, random key in production

def _wants_json_response():
    try:
        accept = request.headers.get('Accept', '') or ''
    except Exception:
        accept = ''
    # Treat any /api/* path or explicit JSON Accept header as API that expects JSON
    return request.path.startswith('/api/') or ('application/json' in accept)

def login_required(view):
    @functools.wraps(view)
    def wrapped_view(**kwargs):
        if 'logged_in' not in session:
            if _wants_json_response():
                return jsonify({'status': 'unauthenticated', 'message': 'Login required'}), 401
            return redirect(url_for('login'))
        return view(**kwargs)
    return wrapped_view

@app.route('/')
def serve_index():
    return send_from_directory(app.static_folder, 'index.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        if username == 'admin' and password == 'certAuth2025#':
            session['logged_in'] = True
            return redirect(url_for('admin_panel'))
        else:
            return render_template('login.html', error='Invalid Credentials')
    return render_template('login.html')

@app.route('/admin')
@login_required
def admin_panel():
    return render_template('admin.html')

@app.route('/data/certs.json')
def serve_certs_json():
    return send_from_directory(os.path.join(app.static_folder, 'data'), 'certs.json')

@app.route('/api/add-certificate', methods=['POST'])
@login_required
def add_certificate():
    try:
        data = request.json
        # Generate a unique CertID (e.g., using a timestamp or UUID)
        cert_id = str(uuid.uuid4()) # Using UUID for uniqueness
        recipient_name = data['recipientName']
        course_title = data['courseTitle']
        date_issued = data['dateIssued']

        # Path to Certificates.csv
        csv_path = os.path.join(os.path.dirname(__file__), '..', 'data', 'Certificates.csv')

        # Read existing CSV and append new data
        with open(csv_path, 'a', newline='') as csvfile:
            fieldnames = ['CertID', 'RecipientName', 'CourseTitle', 'DateIssued']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            # Check if file is empty to write header
            if os.stat(csv_path).st_size == 0:
                writer.writeheader()

            writer.writerow({
                'CertID': cert_id,
                'RecipientName': recipient_name,
                'CourseTitle': course_title,
                'DateIssued': date_issued
            })
        
        # Regenerate all certificate data and QR codes
        vercel_domain = os.environ.get('VERCEL_URL', 'http://localhost:5000')
        project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
        
        input_path = os.path.join(project_root, 'data', 'Certificates.csv')
        output_path = os.path.join(project_root, 'data', 'Certificates_Chained.csv')
        json_path = os.path.join(project_root, 'web', 'data', 'certs.json')
        index_path = os.path.join(project_root, 'web', 'data', 'hash_index.json')
        qr_dir = os.path.join(project_root, 'web', 'data', 'qrcodes')

        chain_certificates.main(
            input_path=input_path,
            output_path=output_path,
            json_path=json_path,
            index_path=index_path,
            base_url=vercel_domain,
            qr_dir=qr_dir,
            add_qr_url=True,
            qr_absolute=True
        )

        return jsonify({'message': 'Certificate added and data regenerated successfully!'}), 200
    except Exception as e:
        import traceback
        app.logger.error("Error adding certificate: %s", traceback.format_exc())
        return jsonify({'message': f'Error adding certificate: {str(e)}'}), 500

@app.route('/api/regenerate-data', methods=['POST'])
@login_required
def regenerate_data():
    try:
        # Assuming the Vercel domain is passed or configured
        # For now, let's use a placeholder or retrieve from environment
        # In a real Vercel deployment, you'd get this from an env var
        vercel_domain = os.environ.get('VERCEL_URL', 'http://localhost:5000') # Fallback for local testing
        
        project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

        input_path = os.path.join(project_root, 'data', 'Certificates.csv')
        output_path = os.path.join(project_root, 'data', 'Certificates_Chained.csv')
        json_path = os.path.join(project_root, 'web', 'data', 'certs.json')
        index_path = os.path.join(project_root, 'web', 'data', 'hash_index.json')
        qr_dir = os.path.join(project_root, 'web', 'data', 'qrcodes')

        chain_certificates.main(
            base_path=project_root,
            input_path=os.path.join(project_root, 'data', 'Certificates.csv'),
            output_path=os.path.join(project_root, 'data', 'Certificates_Chained.csv'),
            json_path=os.path.join(project_root, 'web', 'data', 'certs.json'),
            index_path=os.path.join(project_root, 'web', 'data', 'hash_index.json'),
            base_url=request.url_root.replace('http://', 'https://'),
            qr_dir=os.path.join(project_root, 'web', 'data', 'qrcodes'),
            qr_absolute=True # Ensure absolute paths for QR codes
        )

        return jsonify({'message': 'Certificate data and QR codes regenerated successfully!'}), 200
    except Exception as e:
        import traceback
        app.logger.error("Error regenerating data: %s", traceback.format_exc())
        return jsonify({'message': f'Error regenerating data: {str(e)}'}), 500

@app.route('/api/status')
def api_status():
    if 'logged_in' in session and session['logged_in']:
        return jsonify({'status': 'authenticated'}), 200
    else:
        return jsonify({'status': 'unauthenticated'}), 401

@app.route('/logout', methods=['POST'])
def logout():
    session.pop('logged_in', None)
    return redirect(url_for('login'))

@app.route('/api/regenerate-pdfs', methods=['POST'])
@login_required
def regenerate_pdfs():
    try:
        # Call the generate_certificates_pdf script
        # subprocess.run([sys.executable, os.path.join(app.root_path, '..', 'scripts', 'generate_certificates_pdf.py')], check=True)
        generate_certificates_pdf.generate_all()
        return jsonify({'message': 'Certificate PDFs regenerated successfully!'}), 200
    except Exception as e:
        return jsonify({'message': f'Error regenerating PDFs: {str(e)}'}), 500

@app.route('/api/regenerate-mail-merge', methods=['POST'])
@login_required
def regenerate_mail_merge():
    try:
        project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
        input_csv_path = os.path.join(project_root, 'data', 'Certificates_Chained.csv')
        template_path = os.path.join(project_root, 'templates', 'certificate_template.docx')
        output_docx_path = os.path.join(project_root, 'output', 'Certificates_Merged.docx')
        # Call the mail-merge script directly (original working behavior)
        export_merge_docx.main(
            input_csv_path=input_csv_path,
            template_path=template_path,
            output_docx_path=output_docx_path,
            base_path=project_root
        )
        return jsonify({'message': 'Mail merge DOCX regenerated successfully!'}), 200
    except Exception as e:
        import traceback
        app.logger.error("Error regenerating mail merge DOCX: %s", traceback.format_exc())
        return jsonify({'message': f'Error regenerating mail merge DOCX: {str(e)}'}), 500

@app.route('/api/download-mail-merge')
@login_required
def download_mail_merge():
    try:
        project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
        output_dir = os.path.join(project_root, 'output')
        return send_from_directory(output_dir, 'Certificates_Merged.docx', as_attachment=True)
    except Exception as e:
        import traceback
        app.logger.error("Error downloading mail merge DOCX: %s", traceback.format_exc())
        return jsonify({'message': f'Error downloading mail merge DOCX: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True)