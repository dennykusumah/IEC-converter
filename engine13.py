"""
engine13.py – PDF to DOCX Converter
Mendukung PDF digital dan PDF scan (OCR via Tesseract).
"""

import os
import fitz  # PyMuPDF
from pdf2docx import Converter

class PDFConverterEngine:

    def __init__(self, tesseract_path=None):
        self.tesseract_path = tesseract_path
        if tesseract_path and os.path.exists(tesseract_path):
            try:
                import pytesseract
                pytesseract.pytesseract.tesseract_cmd = tesseract_path
            except ImportError:
                print("Warning: pytesseract not installed. OCR will not work.")

    def is_scanned_pdf(self, pdf_path: str) -> bool:
        try:
            doc = fitz.open(pdf_path)
            text_found = False
            for page in doc:
                if page.get_text().strip():
                    text_found = True
                    break
            doc.close()
            return not text_found
        except Exception:
            return False

    def convert(self, pdf_path: str, docx_path: str) -> tuple:
        try:
            is_scan = self.is_scanned_pdf(pdf_path)
            cv = Converter(pdf_path)
            
            if is_scan:
                mode = "PDF Scan (OCR Active)"
                cv.convert(docx_path, start=0, end=None, ocr=True)
            else:
                mode = "PDF Digital"
                cv.convert(docx_path, start=0, end=None)

            cv.close()
            return True, docx_path, mode
        except Exception as e:
            return False, str(e), "Unknown"