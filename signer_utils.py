import pdfplumber
import re
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PyPDF2 import PdfReader, PdfWriter
import io
from docx import Document
from PIL import Image
import easyocr

EMAIL_RE = re.compile(r'[\w\.-]+@[\w\.-]+\.\w+')

# Initialize EasyOCR reader (do this once globally to avoid reloading model each time)
ocr_reader = None

def get_ocr_reader():
    global ocr_reader
    if ocr_reader is None:
        ocr_reader = easyocr.Reader(['en'])
    return ocr_reader

# ---------------- PDF Extractor ----------------
def extract_emails_and_sigboxes_from_pdf(pdf_path):
    emails = []
    sig_boxes = []
    with pdfplumber.open(pdf_path) as pdf:
        for pageno, page in enumerate(pdf.pages):
            words = page.extract_words()
            for w in words:
                txt = w.get("text", "")
                m = EMAIL_RE.findall(txt)
                if m:
                    for em in m:
                        if em not in emails:
                            emails.append(em)

                low = txt.lower()
                if any(k in low for k in ("sign here", "signature", "signed by", "sign:")):
                    x0 = float(w['x0']) - 10
                    x1 = float(w['x1']) + 150
                    top = float(w['top']) - 5
                    bottom = float(w['bottom']) + 40
                    sig_boxes.append({"page": pageno, "x0": x0, "top": top, "x1": x1, "bottom": bottom})
    return emails, sig_boxes


# ---------------- DOCX Extractor ----------------
def extract_emails_and_sigpos_from_docx(docx_path):
    doc = Document(docx_path)
    emails = []
    sig_positions = []
    for i, p in enumerate(doc.paragraphs):
        text = p.text or ""
        for em in EMAIL_RE.findall(text):
            if em not in emails:
                emails.append(em)
        if "signature" in text.lower() or "[[SIGN_HERE]]" in text:
            sig_positions.append({"para_index": i, "text": text})
    return emails, sig_positions


# ---------------- Image Extractor ----------------
def extract_emails_from_image(image_path):
    """Extract emails from images using EasyOCR."""
    try:
        reader = get_ocr_reader()
        # Read text from image (detail=0 returns just the text)
        results = reader.readtext(image_path, detail=0)
        extracted_text = " ".join(results)
        print(f"DEBUG - Extracted text from {image_path}: {extracted_text}")
        
        # Find emails in the extracted text
        emails = EMAIL_RE.findall(extracted_text)
        print(f"DEBUG - Found emails: {emails}")
        
        # For images, we return None for sig_boxes since we can't precisely locate signature areas
        return emails, None
    except Exception as e:
        print(f"OCR error for {image_path}: {e}")
        return [], None


# ---------------- Overlay Signature ----------------
def overlay_signature_on_pdf(pdf_path, signature_img_path):
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        words = page.extract_words()
        target = None
        for w in words:
            low = w.get("text", "").lower()
            if any(k in low for k in ("sign here", "signature", "sign:", "signed by")):
                x0 = float(w['x0'])
                x1 = float(w['x1'])
                top = float(w['top'])
                bottom = float(w['bottom'])
                target = {"page": 0, "x0": x0, "x1": x1, "top": top, "bottom": bottom}
                break
        if not target:
            target = {"page": 0, "x0": page.width - 200, "x1": page.width - 20,
                      "top": page.height - 200, "bottom": page.height - 100}

    reader = PdfReader(pdf_path)
    writer = PdfWriter()

    for i, p in enumerate(reader.pages):
        packet = io.BytesIO()
        media = p.mediabox
        pw = float(media.width)
        ph = float(media.height)
        can = canvas.Canvas(packet, pagesize=(pw, ph))
        if i == target["page"]:
            y = ph - target["bottom"]
            x = target["x0"]
            sig_width = target["x1"] - target["x0"]
            sig_height = sig_width * 0.4
            can.drawImage(signature_img_path, x, y, width=sig_width, height=sig_height, mask='auto')
        can.save()
        packet.seek(0)
        overlay_pdf = PdfReader(packet)
        base_page = p
        base_page.merge_page(overlay_pdf.pages[0])
        writer.add_page(base_page)

    out_stream = io.BytesIO()
    writer.write(out_stream)
    out_stream.seek(0)
    return out_stream