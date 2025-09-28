import os
import logging
from datetime import datetime
import shutil
from werkzeug.utils import secure_filename
from flask import Flask, request, render_template, redirect, url_for, send_file, flash
from itsdangerous import URLSafeSerializer
from dotenv import load_dotenv
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail
from PyPDF2 import PdfReader

# Import the newly consolidated utility functions
from signer_utils import (
    extract_emails_and_sigboxes_from_pdf,
    extract_emails_and_sigpos_from_docx,
    extract_emails_from_image,
    find_signature_candidates_by_ocr,
    overlay_signature_on_pdf_at_candidates,
    convert_image_to_pdf
)

# Load environment variables from .env file
load_dotenv()

# Basic logging configuration
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "a-very-secret-key-that-you-should-change")

# --- Configuration and Setup ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
SIGNED_FOLDER = os.path.join(BASE_DIR, "signed")
STATIC_FOLDER = os.path.join(BASE_DIR, "static")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(SIGNED_FOLDER, exist_ok=True)

# Email config (using SendGrid)
FROM_EMAIL = os.environ.get("FROM_EMAIL")
SENDGRID_API_KEY = os.environ.get("SENDGRID_API_KEY")

# Token serializer for secure links
serializer = URLSafeSerializer(app.secret_key)

# Allowed upload extensions
ALLOWED_EXT = {'.pdf', '.docx', '.doc', '.png', '.jpg', '.jpeg', '.tiff', '.bmp'}

def allowed_filename(fname: str) -> bool:
    """Checks if the uploaded file extension is allowed."""
    _, ext = os.path.splitext(fname.lower())
    return ext in ALLOWED_EXT

# --- Core Functions ---

def send_sign_email(to_email, token):
    """Sends a signing request email via SendGrid."""
    if not SENDGRID_API_KEY or not FROM_EMAIL:
        logger.warning("SENDGRID_API_KEY or FROM_EMAIL not set. Skipping email.")
        return False
        
    link = url_for('preview_document', token=token, _external=True)
    message = Mail(
        from_email=FROM_EMAIL,
        to_emails=to_email,
        subject="Action Required: Please Sign Your Document",
        html_content=f"""
        <p>You have been requested to sign a document.</p>
        <p>Please click the link below to review and sign:</p>
        <p><a href="{link}" style="padding: 10px 15px; background-color: #007bff; color: white; text-decoration: none; border-radius: 5px;">Review & Sign Document</a></p>
        <p>If you cannot click the link, copy and paste this URL into your browser:</p>
        <p>{link}</p>
        """
    )
    try:
        sg = SendGridAPIClient(SENDGRID_API_KEY)
        response = sg.send(message)
        logger.info(f"Email sent to {to_email}, status code: {response.status_code}")
        return True
    except Exception as e:
        logger.exception(f"Failed to send email to {to_email} via SendGrid: {e}")
        return False

# --- Flask Routes ---

@app.route('/', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        f = request.files.get('file')
        if not f or f.filename == "":
            flash("No file selected. Please choose a file to upload.", "warning")
            return redirect(request.url)

        filename = secure_filename(f.filename)
        if not allowed_filename(filename):
            flash("Unsupported file type. Please upload a PDF, DOCX, or image file.", "danger")
            return redirect(request.url)

        save_path = os.path.join(UPLOAD_FOLDER, filename)
        f.save(save_path)
        logger.info(f"Saved uploaded file to {save_path}")

        # If user uploaded an image, convert it to PDF for a consistent workflow
        _, ext = os.path.splitext(filename.lower())
        if ext in {'.png', '.jpg', '.jpeg', '.tiff', '.bmp'}:
            try:
                converted_pdf_path = convert_image_to_pdf(save_path)
                filename = os.path.basename(converted_pdf_path)
                save_path = converted_pdf_path  # Update path to point to the new PDF
                logger.info(f"Converted image to PDF: {save_path}")
            except Exception as e:
                flash(f"Failed to convert image to PDF: {e}", "danger")
                return redirect(request.url)

        # Extract emails from the document (now guaranteed to be PDF or DOCX)
        emails = []
        lower_filename = filename.lower()
        if lower_filename.endswith('.pdf'):
            emails, _ = extract_emails_and_sigboxes_from_pdf(save_path)
        elif lower_filename.endswith(('.docx', '.doc')):
            emails, _ = extract_emails_and_sigpos_from_docx(save_path)
        
        if not emails:
            flash("No signer email addresses found within the document.", "danger")
            return redirect(request.url)

        # Send a unique signing link to each found email address
        for email in emails:
            payload = {"filename": filename, "email": email}
            token = serializer.dumps(payload)
            send_sign_email(email, token)

        flash(f"Signing requests have been sent to: {', '.join(emails)}", "success")
        return redirect(url_for('upload'))

    return render_template("upload.html")


@app.route('/preview/<token>')
def preview_document(token):
    """Displays the document preview page to the user who clicked the email link."""
    try:
        data = serializer.loads(token)
    except Exception:
        return "The signing link is invalid or has expired.", 400

    filename = data['filename']
    file_url = url_for('uploaded_file', filename=filename)
    return render_template("preview.html",
                           file_url=file_url,
                           token=token,
                           email=data.get('email'))

@app.route('/sign/<token>', methods=['POST'])
def sign_document(token):
    """Handles the final signing process triggered from the preview page."""
    try:
        data = serializer.loads(token)
    except Exception:
        return "The signing link is invalid or has expired.", 400

    filename = data['filename']
    pdf_path = os.path.join(UPLOAD_FOLDER, filename)
    sig_image_path = os.path.join(STATIC_FOLDER, "signature.png")

    if not os.path.exists(pdf_path):
        return "The original document could not be found.", 404
    if not os.path.exists(sig_image_path):
        return "The signature image is missing on the server.", 500

    # Step 1: Use advanced OCR to find the best signature locations.
    try:
        candidates = find_signature_candidates_by_ocr(pdf_path)
    except Exception as e:
        logger.exception(f"OCR candidate detection failed for {filename}: {e}")
        candidates = []

    # Step 2: If OCR fails, fall back to a simpler text-based heuristic.
    if not candidates:
        logger.info(f"OCR found no candidates for {filename}. Falling back to text search.")
        try:
            _, sig_boxes = extract_emails_and_sigboxes_from_pdf(pdf_path)
            if sig_boxes:
                reader = PdfReader(pdf_path)
                for s in sig_boxes:
                    page_num = s.get('page', 0)
                    page = reader.pages[page_num]
                    page_height = float(page.mediabox.height)
                    # Convert top-down coordinates to bottom-up for PDF canvas
                    y_coord = page_height - s.get('y1_bottom', 90)
                    candidates.append({
                        "page": page_num, "x": s.get('x0', 50), "y": y_coord,
                        "width": max(150, s.get('x1', 200) - s.get('x0', 50)),
                        "height": max(40, s.get('y1_bottom', 90) - s.get('y0_top', 50)),
                        "score": 0.4, "reason": "text_heuristic"
                    })
        except Exception as e:
            logger.exception(f"Fallback text heuristic failed for {filename}: {e}")

    if not candidates:
        return "Could not automatically find a suitable place for the signature in the document.", 400

    # Step 3: Overlay the signature onto the best candidate location.
    try:
        out_stream = overlay_signature_on_pdf_at_candidates(pdf_path, sig_image_path, candidates, pick_first=True)
    except Exception as e:
        logger.exception(f"Failed during signature overlay for {filename}: {e}")
        return f"An error occurred while applying the signature.", 500

    # Step 4: Save the newly signed document and send it to the user for download.
    signed_name = f"signed_{datetime.utcnow().strftime('%Y%m%d%H%M%S')}_{filename}"
    out_path = os.path.join(SIGNED_FOLDER, signed_name)
    with open(out_path, "wb") as f:
        f.write(out_stream.read())

    return send_file(out_path, as_attachment=True)


@app.route('/uploads/<filename>')
def uploaded_file(filename):
    """Serves the original uploaded files for previewing."""
    return send_file(os.path.join(UPLOAD_FOLDER, filename))


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)