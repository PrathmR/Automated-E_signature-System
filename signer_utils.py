import pdfplumber
import re
from reportlab.pdfgen import canvas
from PyPDF2 import PdfReader, PdfWriter
import io
from docx import Document
import easyocr
import logging

logger = logging.getLogger(__name__)

EMAIL_RE = re.compile(r'[\w\.-]+@[\w\.-]+\.\w+')

# Initialize EasyOCR reader lazily
ocr_reader = None
def get_ocr_reader():
    global ocr_reader
    if ocr_reader is None:
        # keep language minimal first; expand if needed
        ocr_reader = easyocr.Reader(['en'])
    return ocr_reader

# ---------------- PDF Extractor ----------------
def extract_emails_and_sigboxes_from_pdf(pdf_path):
    """
    Return (emails, sig_boxes)
    sig_boxes is a list of dicts: {"page": int, "x0": float, "x1": float, "top": float, "bottom": float}
    """
    emails = []
    sig_boxes = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for pageno, page in enumerate(pdf.pages):
                words = page.extract_words()
                for w in words:
                    txt = w.get("text", "")
                    for em in EMAIL_RE.findall(txt):
                        if em not in emails:
                            emails.append(em)

                    low = txt.lower()
                    if any(k in low for k in ("sign here", "signature", "signed by", "sign:")):
                        x0 = float(w['x0']) - 10
                        x1 = float(w['x1']) + 150
                        top = float(w['top']) - 5
                        bottom = float(w['bottom']) + 40
                        sig_boxes.append({"page": pageno, "x0": x0, "x1": x1, "top": top, "bottom": bottom})
    except Exception as e:
        logger.exception("Failed to extract from PDF %s: %s", pdf_path, e)
    return emails, sig_boxes

# ---------------- DOCX Extractor ----------------
def extract_emails_and_sigpos_from_docx(docx_path):
    """
    Return (emails, sig_positions)
    sig_positions: simple list describing where signature text occurs in doc (paragraph index + text)
    """
    emails = []
    sig_positions = []
    try:
        doc = Document(docx_path)
        for i, p in enumerate(doc.paragraphs):
            text = p.text or ""
            for em in EMAIL_RE.findall(text):
                if em not in emails:
                    emails.append(em)
            if "signature" in text.lower() or "[[SIGN_HERE]]" in text:
                sig_positions.append({"para_index": i, "text": text})
    except Exception as e:
        logger.exception("Failed to extract from DOCX %s: %s", docx_path, e)
    return emails, sig_positions

# ---------------- Image Extractor ----------------
def extract_emails_from_image(image_path):
    """
    Extract emails using OCR. Returns (emails, None) for sig_boxes since location isn't precise.
    """
    try:
        reader = get_ocr_reader()
        results = reader.readtext(image_path, detail=0)
        extracted_text = " ".join(results)
        emails = EMAIL_RE.findall(extracted_text)
        return emails, None
    except Exception as e:
        logger.exception("OCR error for %s: %s", image_path, e)
        return [], None

# ---------------- Overlay Signature ----------------
def overlay_signature_on_pdf(pdf_path, signature_img_path, sig_boxes=None):
    """
    Overlay signature image onto the PDF.
    If sig_boxes is provided (list), use the first matching box for placement.
    Otherwise try to detect a 'sign' keyword on the first page; if not found, use a default bottom-right placement.
    Returns a BytesIO stream containing the signed PDF.
    """
    try:
        # Determine a target placement if sig_boxes provided
        target = None
        if sig_boxes:
            # Prefer a box on page 0 if present, otherwise take first
            for b in sig_boxes:
                if b.get("page", 0) == 0:
                    target = {"page": 0, "x0": b["x0"], "x1": b["x1"], "top": b["top"], "bottom": b["bottom"]}
                    break
            if target is None and len(sig_boxes) > 0:
                b = sig_boxes[0]
                target = {"page": b.get("page", 0), "x0": b["x0"], "x1": b["x1"], "top": b["top"], "bottom": b["bottom"]}

        # If no target found yet, try to locate text on page 0
        if target is None:
            with pdfplumber.open(pdf_path) as pdf:
                page = pdf.pages[0]
                words = page.extract_words()
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
                    # default bottom-right placement
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
                # compute coordinates
                y = ph - target["bottom"]
                x = target["x0"]
                sig_width = max(50, target["x1"] - target["x0"])
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

    except Exception as e:
        logger.exception("Failed to overlay signature on %s: %s", pdf_path, e)
        # In case of failure, return the original PDF bytes as fallback
        fallback = io.BytesIO()
        try:
            with open(pdf_path, "rb") as fh:
                fallback.write(fh.read())
            fallback.seek(0)
            return fallback
        except Exception:
            # ultimate fallback: empty stream
            logger.exception("Failed to create fallback stream for %s", pdf_path)
            return io.BytesIO()
