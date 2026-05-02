"""
engine12.py – IEC Footer Text Cleaner (Akurat untuk Teks Body/Canvas)
Menghapus teks spesifik di seluruh halaman berdasarkan keselarasan horizontal (Posisi Y):
  - Nomor halaman (contoh: –7–)
  - Nomor standar IEC/ISO (contoh: IEC 62499:2021)
  - Copyright IEC/ISO (contoh: © IEC 2021)
  - Garis horizontal tipis pembatas footer

Cara Kerja: Mencari posisi Y nomor halaman, lalu menghapus semua teks IEC/Copyright 
yang berada di garis Y yang sama persis (toleransi 5 poin).
"""

import fitz  # PyMuPDF
import re
import logging
from typing import Tuple, Any, Dict

logger = logging.getLogger(__name__)


class TextCleanerEngine:
    def __init__(self, bg_color=(1, 1, 1)):
        """
        Args:
            bg_color: Warna untuk menutupi teks yang dihapus. 
                      Default (1, 1, 1) adalah Putih.
        """
        self.bg_color = bg_color

    def clean_iec_footer(self, input_path: str, output_path: str = None) -> Tuple[bool, Any, Dict]:
        try:
            doc = fitz.open(input_path)
        except Exception as e:
            logger.error(f"Gagal membuka PDF: {e}")
            return False, str(e), {}

        total_instances_removed = 0
        y_tolerance = 5  # Toleransi perbedaan posisi Y dalam poin (~1.7mm) untuk menangkap teks sejajar

        # Pola regex
        page_num_pattern = re.compile(r"[–\-]\s*\d+\s*[–\-]")
        iec_std_pattern = re.compile(r"\b(?:IEC|ISO)\s+\d+(?:\-\d+)?:\d{4}\b")
        copyright_pattern = re.compile(r"©\s*(?:IEC|ISO)\s+\d{4}")

        for page_num, page in enumerate(doc):
            areas_to_redact = []
            text_dict = page.get_text("dict")
            footer_y_center = None

            # ── TAHAP 1: Cari posisi Y dari nomor halaman ───────────────
            for block in text_dict["blocks"]:
                if "lines" not in block:
                    continue
                for line in block["lines"]:
                    for span in line["spans"]:
                        if page_num_pattern.search(span["text"]):
                            bbox = fitz.Rect(span["bbox"])
                            # Hitung titik tengah vertikal dari teks nomor halaman
                            footer_y_center = (bbox.y0 + bbox.y1) / 2
                            break
                    if footer_y_center is not None:
                        break
                if footer_y_center is not None:
                    break

            # ── TAHAP 2: Hapus teks yang sejajar dengan posisi Y tersebut ─
            if footer_y_center is not None:
                for block in text_dict["blocks"]:
                    if "lines" not in block:
                        continue
                    for line in block["lines"]:
                        for span in line["spans"]:
                            bbox = fitz.Rect(span["bbox"])
                            span_y_center = (bbox.y0 + bbox.y1) / 2
                            
                            # Cek apakah teks ini berada di garis yang sama dengan nomor halaman
                            if abs(span_y_center - footer_y_center) <= y_tolerance:
                                text = span["text"]
                                
                                # Hapus jika itu nomor halaman
                                if page_num_pattern.search(text):
                                    areas_to_redact.append(bbox)
                                    total_instances_removed += 1
                                    
                                # Hapus jika itu Standar IEC/ISO atau Copyright
                                elif iec_std_pattern.search(text) or copyright_pattern.search(text):
                                    areas_to_redact.append(bbox)
                                    total_instances_removed += 1

            # ── TAHAP 3: Hapus garis horizontal tipis di atas footer ────
            # (Dokumen IEC biasanya punya garis tipis pembatas footer)
            if footer_y_center is not None:
                paths = page.get_drawings()
                for p in paths:
                    rect = p["rect"]
                    # Identifikasi garis horizontal (tinggi sangat kecil, lebar memanjang)
                    if rect.height < 2 and rect.width > 50:
                        line_y_center = (rect.y0 + rect.y1) / 2
                        # Jika garis berada di dekat area footer (di atasnya)
                        if abs(line_y_center - footer_y_center) <= 15:
                            page.add_redact_annot(rect, fill=self.bg_color)

            # ── TAHAP 4: Terapkan Redaction ─────────────────────────────
            for area in areas_to_redact:
                # Padding 2pt agar teks tertutup rapi tanpa sisa
                padded_area = area + (-2, -2, 2, 2) 
                page.add_redact_annot(padded_area, fill=self.bg_color)
            
            if areas_to_redact:
                page.apply_redactions()
                logger.debug(f"Halaman {page_num + 1}: {len(areas_to_redact)} elemen footer dihapus berdasarkan posisi Y.")

        info = {
            "total_instances_removed": total_instances_removed
        }

        if output_path:
            doc.save(output_path)
            doc.close()
            return True, output_path, info
        else:
            pdf_bytes = doc.tobytes()
            doc.close()
            return True, pdf_bytes, info
