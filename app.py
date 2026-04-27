"""
app.py – IEC/ISO PDF Ultimate Processor
Alur:
  1. Trim PDF (engine11.py) -> Potong bahasa Prancis
  2. Clean Footer (engine12.py) -> Hapus nomor halaman & copyright IEC
  3. Convert to DOCX (engine13.py) -> Siap diterjemahkan
"""

import os
import streamlit as st
import logging
from engine11 import PDFTrimmerEngine
from engine12 import TextCleanerEngine
from engine13 import PDFConverterEngine

# ── Setup Logging ─────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO)

# ── Konfigurasi Halaman ────────────────────────────────────────────────────
st.set_page_config(page_title="IEC/ISO PDF Processor", page_icon="📄", layout="centered")

# ── Inisialisasi Engine ────────────────────────────────────────────────────
trimmer = PDFTrimmerEngine()
cleaner = TextCleanerEngine(bg_color=(1, 1, 1))

DEFAULT_TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
converter = PDFConverterEngine(tesseract_path=DEFAULT_TESSERACT_PATH)

# ── Buat Folder Temp ──────────────────────────────────────────────────────
if not os.path.exists("temp"):
    os.makedirs("temp")

# ── UI Utama ───────────────────────────────────────────────────────────────
st.title("📄 IEC/ISO PDF Processor")

st.divider()

uploaded_file = st.file_uploader("Pilih file PDF", type=["pdf"])

if uploaded_file:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        process_btn = st.button("🚀 Jalankan Semua Proses", type="primary", use_container_width=True)

    if process_btn:
        base_filename = os.path.splitext(uploaded_file.name)[0]
        
        input_pdf = os.path.join("temp", uploaded_file.name)
        trimmed_pdf = os.path.join("temp", f"{base_filename}_1_trimmed.pdf")
        cleaned_pdf = os.path.join("temp", f"{base_filename}_2_cleaned.pdf")
        output_docx = os.path.join("temp", f"{base_filename}_3_result.docx")

        with open(input_pdf, "wb") as f:
            f.write(uploaded_file.getbuffer())

        with st.status("Memproses...", expanded=True) as status:
            
            try:
                # ── TAHAP 1: TRIMMING ─────────────────────────────────
                st.write("✂️ Tahap 1...")
                success_trim, res_trim, info_trim = trimmer.trim(input_path=input_pdf, output_path=trimmed_pdf)
                
                if not success_trim:
                    st.error(res_trim)
                    status.update(label="Gagal", state="error")
                else:
                    st.write("✔️ Selesai")
                    
                    # ── TAHAP 2: CLEANING FOOTER ──────────────────────
                    st.write("🧹 Tahap 2...")
                    success_clean, res_clean, info_clean = cleaner.clean_iec_footer(input_path=trimmed_pdf, output_path=cleaned_pdf)
                    
                    if not success_clean:
                        st.error(res_clean)
                        status.update(label="Gagal", state="error")
                    else:
                        st.write("✔️ Selesai")
                        
                        # ── TAHAP 3: CONVERSION ──────────────────────
                        st.write("🔄 Tahap 3...")
                        success_conv, res_conv, mode_conv = converter.convert(cleaned_pdf, output_docx)
                        
                        if not success_conv:
                            st.error(res_conv)
                            status.update(label="Gagal", state="error")
                        else:
                            st.write("✔️ Selesai")
                            
                            with open(output_docx, "rb") as f:
                                st.session_state["docx_bytes"] = f.read()
                                
                            st.session_state["download_filename"] = f"{base_filename}_RESULT.docx"
                            status.update(label="✅ Selesai", state="complete", expanded=False)
                            
            except Exception as e:
                st.error(str(e))
                status.update(label="Error", state="error")
                
            finally:
                for path in [input_pdf, trimmed_pdf, cleaned_pdf, output_docx]:
                    try:
                        if path and os.path.exists(path): os.remove(path)
                    except Exception:
                        pass

# ── Panel Download ─────────────────────────────────────────────────────────
if "docx_bytes" in st.session_state:
    st.divider()
    
    st.download_button(
        label="⬇️ Download .DOCX",
        data=st.session_state["docx_bytes"],
        file_name=st.session_state["download_filename"],
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
