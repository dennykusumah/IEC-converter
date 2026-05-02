"""
engine11.py – PDF Trimmer Engine untuk Standar IEC/ISO (FIXED BILINGUAL)

Memotong halaman PDF:
  - Menghapus halaman dari halaman 1 sampai sebelum halaman Scope/Ruang Lingkup
  - Menghapus halaman dari SETELAH Bibliography (Inggris) sampai halaman terakhir
  - Efek: Seluruh bagian bahasa Prancis di belakang akan terhapus otomatis.

Dependency: PyMuPDF  →  pip install PyMuPDF
"""

import fitz  # PyMuPDF
import re
import logging
from typing import Tuple, Optional, Dict, Any

logger = logging.getLogger(__name__)


class PDFTrimmerEngine:
    """
    Engine untuk memotong (trim) halaman PDF standar IEC/ISO.
    """

    # ── Pola regex untuk mendeteksi heading Scope / Ruang Lingkup ──────────
    SCOPE_PATTERNS = [
        r"(?i)\b1[\.\s]+scope\b",               # "1 Scope", "1. Scope"
        r"(?i)\b1[\.\s]+ruang\s+lingkup\b",     # versi Indonesia
        r"(?i)\b1\s*scope\b",                   # "1Scope"
    ]

    # ── Pola regex untuk mendeteksi Bibliography (DIPERBAIKI UNTUK BILINGUAL) ─
    # Menggunakan \b (word boundary) di akhir agar TIDAK cocok dengan "Bibliographie" (Prancis)
    BIB_PATTERNS = [
        r"(?i)\bbibliography\b",                # Hanya mencocokkan "Bibliography" (Inggris)
        r"(?i)\bbibliografi\b",                 # Hanya mencocokkan "Bibliografi" (Indonesia)
    ]

    # ─────────────────────────────────────────────────────────────────────
    # Internal helpers
    # ─────────────────────────────────────────────────────────────────────

    @staticmethod
    def _is_toc_page(page: fitz.Page, threshold: int = 5) -> bool:
        """Heuristik sederhana untuk mendeteksi halaman Daftar Isi (ToC)."""
        text = page.get_text()
        dot_leaders = len(re.findall(r"\.{3,}", text))
        lines_ending_num = len(re.findall(r"\.{2,}\s*\d+\s*$", text, re.MULTILINE))
        return dot_leaders > threshold or lines_ending_num > threshold

    # ─────────────────────────────────────────────────────────────────────
    # Pencarian halaman kunci
    # ─────────────────────────────────────────────────────────────────────

    def find_scope_page(self, doc: fitz.Document) -> Optional[int]:
        """Mencari indeks halaman (0-based) yang berisi heading 'Scope'."""
        for page_num in range(len(doc)):
            page = doc[page_num]
            if self._is_toc_page(page):
                logger.debug(f"Melewatkan halaman {page_num + 1} (terdeteksi ToC)")
                continue
            
            text = page.get_text()
            # Gabungkan semua baris menjadi 1 baris panjang untuk menghindari kegagalan regex
            clean_text = re.sub(r"\s+", " ", text) 
            
            for pattern in self.SCOPE_PATTERNS:
                if re.search(pattern, clean_text):
                    logger.info(f"Scope ditemukan di halaman {page_num + 1}")
                    return page_num
        logger.warning("Halaman Scope tidak ditemukan")
        return None

    def find_bibliography_page(self, doc: fitz.Document) -> Optional[int]:
        """
        Mencari indeks halaman (0-based) yang berisi heading 'Bibliography'.
        
        PERBAIKAN: Pencarian dilakukan dari DEPAN (bukan dari belakang).
        Ini untuk memastikan yang ditemukan adalah "Bibliography" (Inggris) 
        yang menjadi batas akhir bagian Inggris, BUKAN "Bibliographie" (Prancis)
        yang ada di ujung paling belakang dokumen.
        """
        for page_num in range(len(doc)): # Dari depan ke belakang
            page = doc[page_num]
            
            # Lewati jika ini halaman Daftar Isi (ToC)
            if self._is_toc_page(page):
                continue
                
            text = page.get_text()
            clean_text = re.sub(r"\s+", " ", text)
            
            for pattern in self.BIB_PATTERNS:
                if re.search(pattern, clean_text):
                    logger.info(f"Bibliography (batas akhir bahasa Inggris) ditemukan di halaman {page_num + 1}")
                    return page_num
                    
        logger.warning("Halaman Bibliography tidak ditemukan")
        return None

    # ─────────────────────────────────────────────────────────────────────
    # Method utama: trim
    # ─────────────────────────────────────────────────────────────────────

    def trim(
        self, input_path: str, output_path: str = None
    ) -> Tuple[bool, Any, Dict]:
        """
        Memotong PDF dari halaman Scope sampai Bibliography.
        Halaman Bibliography TETAP ADA, halaman setelahnya (bahasa Prancis) DIHAPUS.
        """
        try:
            doc = fitz.open(input_path)
        except Exception as e:
            logger.error(f"Gagal membuka PDF: {e}")
            return False, str(e), {}

        total_pages = len(doc)
        if total_pages == 0:
            doc.close()
            return False, "PDF kosong (0 halaman).", {}

        # ── Cari halaman kunci ──────────────────────────────────────────
        scope_page = self.find_scope_page(doc)
        bib_page = self.find_bibliography_page(doc)

        start_keep = scope_page if scope_page is not None else 0
        # Jika ditemukan, end_keep = bib_page. Artinya halaman bib_page IKUT disimpan.
        end_keep = bib_page if bib_page is not None else total_pages - 1

        # ── Validasi ────────────────────────────────────────────────────
        if start_keep > end_keep:
            doc.close()
            return False, (
                f"Halaman Scope (hlm. {start_keep + 1}) berada setelah "
                f"Bibliography (hlm. {end_keep + 1}). Tidak dapat memotong."
            ), {}

        pages_removed_before = start_keep
        pages_removed_after = total_pages - end_keep - 1 # Ini adalah jumlah halaman Prancis yang terhapus
        total_trimmed = end_keep - start_keep + 1

        # ── Buat dokumen baru dengan rentang halaman yang dipertahankan ─
        new_doc = fitz.open()
        new_doc.insert_pdf(doc, from_page=start_keep, to_page=end_keep)
        doc.close()

        info = {
            "total_original": total_pages,
            "scope_page": scope_page + 1 if scope_page is not None else None,
            "bib_page": bib_page + 1 if bib_page is not None else None,
            "start_keep": start_keep + 1,       
            "end_keep": end_keep + 1,
            "total_trimmed": total_trimmed,
            "pages_removed_before": pages_removed_before,
            "pages_removed_after": pages_removed_after,
            "scope_found": scope_page is not None,
            "bib_found": bib_page is not None,
        }

        # ── Simpan / kembalikan hasil ───────────────────────────────────
        if output_path:
            new_doc.save(output_path)
            new_doc.close()
            return True, output_path, info
        else:
            pdf_bytes = new_doc.tobytes()
            new_doc.close()
            return True, pdf_bytes, info

    # ─────────────────────────────────────────────────────────────────────
    # Preview (opsional – untuk menampilkan thumbnail di Streamlit)
    # ─────────────────────────────────────────────────────────────────────

    @staticmethod
    def render_page_preview(
        input_path: str, page_num: int, dpi: int = 100
    ) -> Optional[bytes]:
        """Merender satu halaman PDF menjadi gambar PNG."""
        try:
            doc = fitz.open(input_path)
            if page_num < 0 or page_num >= len(doc):
                doc.close()
                return None
            page = doc[page_num]
            pix = page.get_pixmap(dpi=dpi)
            img_bytes = pix.tobytes("png")
            doc.close()
            return img_bytes
        except Exception as e:
            logger.error(f"Gagal render preview halaman {page_num}: {e}")
            return None
