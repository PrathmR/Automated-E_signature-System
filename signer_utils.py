import os
import io
import re
import logging
import difflib
from typing import List, Dict, Tuple, Optional

import cv2
import numpy as np
import pytesseract
import pdfplumber
import easyocr
from pytesseract import Output
from pdf2image import convert_from_path
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from PIL import Image
from docx import Document

# --- Configuration and Setup ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

EMAIL_RE = re.compile(r'[\w\.-]+@[\w\.-]+\.\w+')
KEYWORDS = [
    "signature", "sign here", "sign:", "sig.", "authorised signatory",
    "signed by", "please sign", "customer signature", "employee signature",
    "signatur", "signat", "siganture", "signature:"
]

# Initialize EasyOCR reader lazily for email extraction from images
ocr_reader = None
def get_ocr_reader():
    global ocr_reader
    if ocr_reader is None:
        ocr_reader = easyocr.Reader(['en'])
    return ocr_reader

# --- Helper Functions ---

def fuzzy_match(text: str, keywords=KEYWORDS, cutoff: float = 0.70) -> bool:
    """Fuzzy matches text against a list of signature-related keywords."""
    if not text:
        return False
    t = text.lower()
    for kw in keywords:
        if kw in t or difflib.SequenceMatcher(None, t, kw).ratio() >= cutoff:
            return True
    return False

def safe_ocr_image(pil_img: Image.Image, config: str = "--psm 6") -> dict:
    """Runs Tesseract OCR safely, returning a dictionary of results."""
    try:
        img_cv = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
        data = pytesseract.image_to_data(img_cv, output_type=Output.DICT, config=config)
        # Ensure all expected keys exist in the output
        keys = ["level", "page_num", "block_num", "par_num", "line_num", "word_num",
                "left", "top", "width", "height", "conf", "text"]
        return {k: data.get(k, []) for k in keys}
    except Exception as e:
        logger.exception(f"safe_ocr_image failed: {e}")
        return {k: [] for k in keys}


def convert_ocr_to_pdf_coords(ocr_x, ocr_y, ocr_w, ocr_h, img_w, img_h, pdf_w, pdf_h):
    """Converts OCR pixel coordinates (top-left origin) to PDF points (bottom-left origin)."""
    if img_w == 0 or img_h == 0: return 0, 0, ocr_w, ocr_h
    scale_x = pdf_w / img_w
    scale_y = pdf_h / img_h
    pdf_x = ocr_x * scale_x
    pdf_w_res = ocr_w * scale_x
    pdf_y = pdf_h - ((ocr_y + ocr_h) * scale_y)  # Invert Y-axis
    pdf_h_res = ocr_h * scale_y
    return pdf_x, pdf_y, pdf_w_res, pdf_h_res

def preprocess_image_for_ocr_cv(img_cv: np.ndarray) -> np.ndarray:
    """Applies preprocessing steps to an image to improve OCR accuracy."""
    try:
        gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
        denoised = cv2.medianBlur(gray, 3)
        return cv2.adaptiveThreshold(denoised, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                     cv2.THRESH_BINARY, 11, 2)
    except Exception:
        logger.exception("Image preprocessing failed. Falling back to grayscale.")
        return cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)


def remove_duplicate_candidates(candidates: List[Dict], threshold_px: float = 30.0) -> List[Dict]:
    """Removes nearly overlapping candidate boxes to avoid placing multiple signatures."""
    out = []
    for c in sorted(candidates, key=lambda x: -float(x.get("score", 0))):
        is_duplicate = False
        for u in out:
            if abs(c["x"] - u["x"]) < threshold_px and abs(c["y"] - u["y"]) < threshold_px:
                is_duplicate = True
                break
        if not is_duplicate:
            out.append(c)
    return out

# --- Main Utility Functions ---

def convert_image_to_pdf(image_path: str) -> str:
    """Converts an image file into a single-page PDF."""
    base_dir = os.path.dirname(image_path)
    base_name = os.path.splitext(os.path.basename(image_path))[0]
    out_pdf = os.path.join(base_dir, f"{base_name}_converted.pdf")
    try:
        with Image.open(image_path) as img:
            # Handle transparency
            if img.mode in ("RGBA", "LA"):
                background = Image.new("RGB", img.size, (255, 255, 255))
                background.paste(img, mask=img.split()[-1])
                img = background
            else:
                img = img.convert("RGB")
            img.save(out_pdf, "PDF", resolution=300.0)
        logger.info(f"Converted image {image_path} -> {out_pdf}")
        return out_pdf
    except Exception as e:
        logger.exception(f"convert_image_to_pdf failed for {image_path}")
        raise e

def extract_emails_and_sigboxes_from_pdf(pdf_path: str) -> Tuple[List[str], List[Dict]]:
    """Extracts emails and potential signature boxes from a PDF's text layer."""
    emails, sig_boxes = [], []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for pageno, page in enumerate(pdf.pages):
                full_text = page.extract_text() or ""
                for email in EMAIL_RE.findall(full_text):
                    if email not in emails:
                        emails.append(email)
                
                for w in page.extract_words():
                    if fuzzy_match(w.get("text", "")):
                        sig_boxes.append({
                            "page": pageno,
                            "x0": float(w.get("x0", 0)) - 5,
                            "x1": float(w.get("x1", page.width)),
                            "y0_top": float(w.get("top", 0)) - 4,
                            "y1_bottom": float(w.get("bottom", 0)) + 45
                        })
    except Exception as e:
        logger.exception(f"extract_emails_and_sigboxes_from_pdf failed for {pdf_path}: {e}")
    return list(set(emails)), sig_boxes

def extract_emails_and_sigpos_from_docx(docx_path):
    """Extracts emails and signature keywords from DOCX files."""
    emails, sig_positions = [], []
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
        logger.exception(f"Failed to extract from DOCX {docx_path}: {e}")
    return emails, sig_positions

def extract_emails_from_image(image_path: str) -> Tuple[List[str], None]:
    """Extracts emails from an image file using EasyOCR."""
    try:
        reader = get_ocr_reader()
        results = reader.readtext(image_path, detail=0)
        emails = EMAIL_RE.findall(" ".join(results))
        return list(set(emails)), None
    except Exception as e:
        logger.exception(f"OCR email extraction error for {image_path}: {e}")
        return [], None

def find_signature_candidates_by_ocr(pdf_path: str, dpi: int = 300) -> List[Dict]:
    """
    Finds signature locations and defines a standard, larger box for legibility.
    """
    candidates = []
    try:
        pil_pages = convert_from_path(pdf_path, dpi=dpi)
        reader = PdfReader(pdf_path)
    except Exception as e:
        logger.exception(f"Failed to open/convert PDF: {pdf_path}: {e}")
        return []

    for page_index, pil_img in enumerate(pil_pages):
        page = reader.pages[page_index]
        pdf_w, pdf_h = float(page.mediabox.width), float(page.mediabox.height)
        img_cv = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
        img_h, img_w = img_cv.shape[:2]
        proc_img = preprocess_image_for_ocr_cv(img_cv)

        # OCR Pass
        ocr_data = safe_ocr_image(proc_img)
        for i, text in enumerate(ocr_data["text"]):
            conf = int(float(ocr_data["conf"][i])) if i < len(ocr_data["conf"]) and ocr_data["conf"][i] != '-1' else 0
            if conf > 30 and fuzzy_match(text):
                lx, ty, w_px, h_px = ocr_data["left"][i], ocr_data["top"][i], ocr_data["width"][i], ocr_data["height"][i]
                pdf_x, pdf_y, pdf_w_box, pdf_h_box = convert_ocr_to_pdf_coords(lx, ty, w_px, h_px, img_w, img_h, pdf_w, pdf_h)
                
                # --- NEW LOGIC FOR LARGER SIGNATURE ---
                # Define a practical size for the signature. You can adjust these values.
                signature_box_width = 120  # A good width to cover "Sign Here"
                signature_box_height = 45   # A good legible height

                # Vertically center the new, larger box on the original detected text
                new_box_y = pdf_y + (pdf_h_box / 2) - (signature_box_height / 2)

                candidates.append({
                    "page": page_index,
                    "x": pdf_x,                 # Start where the detected text starts
                    "y": new_box_y,               # Use the new vertically centered 'y'
                    "width": signature_box_width, # Use our new standard width
                    "height": signature_box_height,# Use our new standard height
                    "score": max(0.5, conf / 100.0),
                    "reason": f"ocr: {text[:20]}"
                })
                # --- END NEW LOGIC ---

        # Visual Heuristic: Find long horizontal lines
        try:
            edges = cv2.Canny(proc_img, 50, 150)
            lines = cv2.H_HoughLinesP(edges, 1, np.pi/180, threshold=100, minLineLength=100, maxLineGap=10)
            if lines is not None:
                for ln in lines:
                    x1, y1, x2, y2 = ln[0]
                    if abs(y2 - y1) < 10:
                        pdf_x, pdf_y, line_width, _ = convert_ocr_to_pdf_coords(min(x1,x2), min(y1,y2), abs(x2-x1), 20, img_w, img_h, pdf_w, pdf_h)
                        candidates.append({
                             "page": page_index, "x": pdf_x, "y": pdf_y,
                             "width": line_width, "height": 40, "score": 0.45, "reason": "line_detect"
                        })
        except Exception as e:
            logger.warning(f"Line detection failed on page {page_index}: {e}")

    unique_candidates = remove_duplicate_candidates(candidates)
    logger.info(f"Found {len(unique_candidates)} unique signature candidates in {pdf_path}")
    return sorted(unique_candidates, key=lambda c: c.get("score", 0), reverse=True)
def overlay_signature_on_pdf_at_candidates(pdf_path, signature_img_path, candidates, pick_first=True) -> io.BytesIO:
    """Overlays the signature image, stretching it to fit the candidate box exactly."""
    out_stream = io.BytesIO()
    reader = PdfReader(pdf_path)
    writer = PdfWriter()

    cands_by_page = {}
    for c in candidates:
        cands_by_page.setdefault(c["page"], []).append(c)

    for p_index, page in enumerate(reader.pages):
        packet = io.BytesIO()
        pw, ph = float(page.mediabox.width), float(page.mediabox.height)
        can = canvas.Canvas(packet, pagesize=(pw, ph))

        if p_index in cands_by_page:
            to_draw = [cands_by_page[p_index][0]] if pick_first else cands_by_page[p_index]
            for cand in to_draw:
                # --- CORRECTED LOGIC ---
                # Use the exact position and dimensions from the candidate box.
                # This will stretch/shrink the signature to fit perfectly.
                x = cand["x"]
                y = cand["y"]
                sig_w = cand["width"]
                sig_h = cand["height"]
                # --- END CORRECTION ---

                can.drawImage(signature_img_path, x, y, width=sig_w, height=sig_h, mask='auto')
                logger.info(f"Placed signature on page {p_index} at ({x:.1f}, {y:.1f}) to cover text '{cand['reason']}'")

        can.save()
        packet.seek(0)
        overlay_pdf = PdfReader(packet)
        page.merge_page(overlay_pdf.pages[0])
        writer.add_page(page)

    writer.write(out_stream)
    out_stream.seek(0)
    return out_stream