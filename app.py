import os
import logging
from datetime import datetime
from flask import Flask, request, render_template, redirect, url_for, send_file, flash
from itsdangerous import URLSafeSerializer
from signer_utils import (
    extract_emails_and_sigboxes_from_pdf,
    extract_emails_and_sigpos_from_docx,
    overlay_signature_on_pdf,
    extract_emails_from_image
)
from email.message import EmailMessage
import smtplib
from dotenv import load_dotenv
import shutil

# Load environment variables
load_dotenv()

# Basic logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "change_this_to_a_secure_random_value")

UPLOAD_FOLDER = "uploads"
SIGNED_FOLDER = "signed"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(SIGNED_FOLDER, exist_ok=True)

# Email config
SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER")
SMTP_PASS = os.environ.get("SMTP_PASS")
FROM_EMAIL = os.environ.get("FROM_EMAIL", SMTP_USER)

# Token serializer
serializer = URLSafeSerializer(app.secret_key)

# ---------------- EMAIL FUNCTION ----------------
def send_sign_email(to_email, token):
    """
    Try to send an email with a preview/sign link. Failure does not crash the app;
    it logs the error and returns False on failure.
    """
    if not SMTP_USER or not SMTP_PASS:
        logger.warning("SMTP credentials missing — skipping actual email send.")
        return False

    link = url_for('preview_document', token=token, _external=True)
    msg = EmailMessage()
    msg['Subject'] = 'Please review and sign the document'
    msg['From'] = FROM_EMAIL
    msg['To'] = to_email
    msg.set_content(f"You were requested to sign a document. Review it here: {link}")

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
            s.starttls()
            s.login(SMTP_USER, SMTP_PASS)
            s.send_message(msg)
        logger.info("Sent email to %s", to_email)
        return True
    except Exception as e:
        logger.exception("Failed to send email to %s: %s", to_email, e)
        return False

# ---------------- FILE UPLOAD ----------------
@app.route('/', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        f = request.files.get('file')
        logger.debug("Uploaded file object: %s", f)
        if not f or f.filename == "":
            flash("No file selected")
            return redirect(request.url)

        filename = f.filename
        save_path = os.path.join(UPLOAD_FOLDER, filename)
        f.save(save_path)
        logger.info("Saved uploaded file to %s", save_path)

        emails = []
        sig_boxes = None

        lower = filename.lower()
        if lower.endswith('.pdf'):
            emails, sig_boxes = extract_emails_and_sigboxes_from_pdf(save_path)
        elif lower.endswith(('.docx', '.doc')):
            emails, sig_boxes = extract_emails_and_sigpos_from_docx(save_path)
        elif lower.endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff')):
            emails, sig_boxes = extract_emails_from_image(save_path)
        else:
            flash("Unsupported file type")
            return redirect(request.url)

        if not emails:
            flash("No signer email found in document/image")
            return redirect(request.url)

        candidate = emails[0]
        payload = {"filename": filename, "email": candidate, "sig_boxes": sig_boxes}
        token = serializer.dumps(payload)

        preview_link = url_for('preview_document', token=token, _external=True)
        sent = send_sign_email(candidate, token)
        if sent:
            flash(f"Email sent to {candidate}")
        else:
            flash(f"Could not send email to {candidate}. Preview link: opened locally.")

        return redirect(preview_link)

    return render_template("upload.html")

# ---------------- PREVIEW DOCUMENT ----------------
@app.route('/preview/<token>')
def preview_document(token):
    try:
        data = serializer.loads(token)
    except Exception:
        return "Invalid or expired link", 400

    filename = data['filename']
    file_url = url_for('uploaded_file', filename=filename, _external=True)
    return render_template("preview.html", file_url=file_url, token=token, email=data.get('email'))

# ---------------- SIGN DOCUMENT ----------------
@app.route('/sign/<token>', methods=['POST'])
def sign_document(token):
    try:
        data = serializer.loads(token)
    except Exception:
        return "Invalid or expired link", 400

    filename = data['filename']
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    signed_name = f"signed_{datetime.utcnow().strftime('%Y%m%d%H%M%S')}_{filename}"
    out_path = os.path.join(SIGNED_FOLDER, signed_name)

    # Use detected sig_boxes from payload (may be None)
    sig_boxes = data.get('sig_boxes')

    # Handle PDF files with signature overlay
    if filename.lower().endswith('.pdf'):
        sig_image = os.path.join("static", "signature.png")
        out_stream = overlay_signature_on_pdf(filepath, sig_image, sig_boxes)
        with open(out_path, "wb") as out_f:
            out_f.write(out_stream.read())
    else:
        # For images and other files, just copy them as "signed"
        shutil.copy2(filepath, out_path)

    return send_file(out_path, as_attachment=True)

# ---------------- SERVE UPLOADED FILES ----------------
@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_file(os.path.join(UPLOAD_FOLDER, filename))

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
