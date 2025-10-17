"""
KAY App - FINAL SINGLE PAGE APP (SPA) - VERSI OPTIMALISASI KETERANGAN & UI/UX
- Tujuan: Memastikan struktur kode dan alur aplikasi mudah dimengerti.
- Perbaikan: SyntaxError, penambahan Batch Rename PDF/Gambar Sesuai Excel, dan Reorder PDF.
"""

import os
import io
import zipfile
import shutil
import traceback
import tempfile
import time

import streamlit as st
import pandas as pd
from PIL import Image

# --- 1. SETUP LIBRARY IMPORTS (Error Handling) ---
# PDF libs
PdfReader = PdfWriter = None
try:
    from PyPDF2 import PdfReader, PdfWriter
except Exception:
    pass

pdfplumber = None
try:
    import pdfplumber
except Exception:
    pass

Document = None
try:
    from docx import Document
except Exception:
    pass

# Image libs (pdf2image check)
PDF2IMAGE_AVAILABLE = False
convert_from_bytes = convert_from_path = None
try:
    from pdf2image import convert_from_bytes, convert_from_path 
    PDF2IMAGE_AVAILABLE = True
except Exception:
    try:
        from pdf2image import convert_from_path
        PDF2IMAGE_AVAILABLE = True
        convert_from_bytes = None
    except Exception:
        pass 

# --- 2. FUNGSI HELPER UTAMA ---

def make_zip_from_map(bytes_map: dict) -> bytes:
    """Membuat file ZIP dari kamus {nama_file: data_bytes}."""
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in bytes_map.items():
            z.writestr(name, data)
    b.seek(0)
    return b.getvalue()

def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Mengubah DataFrame Pandas menjadi bytes Excel."""
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    out.seek(0)
    return out.getvalue()

def try_encrypt(writer, password: str):
    """Fungsi aman untuk enkripsi PDF menggunakan PyPDF2."""
    try:
        writer.encrypt(password)
    except Exception:
        writer.encrypt(user_pwd=password, owner_pwd=password)

def navigate_to(target_menu):
    """Helper global untuk navigasi antar halaman/menu menggunakan session state."""
    st.session_state.menu_selection = target_menu
    try:
        st.rerun() 
    except AttributeError:
        # Fallback for older Streamlit versions
        st.experimental_rerun() 

# --- 3. STREAMLIT KONFIGURASI DAN CSS (Tampilan Bersih) ---

LOGO_PATH = os.path.join("assets", "logo.png")
page_icon = LOGO_PATH if os.path.exists(LOGO_PATH) else "🛠️" 
st.set_page_config(page_title="KAY App – Tools MCU", page_icon=page_icon, layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
/* Hilangkan Sidebar dan UI bawaan Streamlit */
[data-testid="stSidebarToggleButton"], 
section[data-testid="stSidebar"],      
[data-testid="stDecoration"],
.stApp a[href*="github.com/"],
.stApp header > div:last-child {
    visibility: hidden !important;
    display: none !important;
    width: 0 !important;
    padding: 0 !important;
}

/* Background gradient lembut */
.stApp {
    background: linear-gradient(180deg, #e9f2ff 0%, #f4f9ff 100%); 
    color: #002b5b;
    font-family: 'Inter', sans-serif;
}

/* Tombol modern dan jelas */
div.stButton > button {
    background: linear-gradient(90deg, #5dade2, #3498db);
    color: white; 
    border: none;
    border-radius: 12px;
    padding: 0.5rem 1rem;
    font-weight: 600;
    transition: 0.2s;
    box-shadow: 0 4px 8px rgba(52, 152, 219, 0.4); 
    cursor: pointer;
    width: 100%; /* Tombol fitur utama dibuat lebar penuh */
}

div.stButton > button:hover {
    background: linear-gradient(90deg, #3498db, #2e86c1); 
    transform: scale(1.01);
    box-shadow: 0 6px 14px rgba(52, 152, 219, 0.5); 
}

/* Card fitur untuk dashboard */
.feature-card {
    background: white;
    border-radius: 12px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06); 
    padding: 18px; 
    transition: 0.2s; 
    border: 1px solid #d0e3ff; 
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    height: 100%; 
}
.feature-card:hover {
    box-shadow: 0 6px 18px rgba(0,0,0,0.10);
    transform: translateY(-2px);
}
h1 { color: #1b4f72; font-weight: 800;}
</style>
""", unsafe_allow_html=True)

# --- 4. INISIALISASI SESSION STATE & HEADER GLOBAL ---

if "menu_selection" not in st.session_state:
    st.session_state.menu_selection = "Dashboard"
    
menu = st.session_state.menu_selection

def add_back_to_dashboard_button():
    """Tombol navigasi kembali yang konsisten di setiap halaman fitur."""
    if st.button("🏠 Kembali ke Dashboard", key="back_to_dash"):
        navigate_to("Dashboard")
    st.markdown("---")

# Header di semua halaman
if os.path.exists(LOGO_PATH):
    try:
        st.image(LOGO_PATH, width=110)
    except Exception:
        st.write("KAY App")
st.title("KAY App – Tools MCU")
st.markdown("Aplikasi serbaguna untuk pengolahan dokumen, PDF, gambar, dan data MCU.")
st.markdown("---")


# =================================================================
# --- 5. LOGIKA NAVIGASI APLIKASI (PER HALAMAN) ---
# =================================================================

# -------------- 5.1 Dashboard (Menu Utama) --------------
if menu == "Dashboard":
    st.markdown("### Pilih Fitur Utama")
    
    cols1 = st.columns(3)

    # Kolom 1: Kompres Foto
    with cols1[0]:
        st.markdown('<div class="feature-card"><b>📸 Kompres & Gambar Tools</b><br>Perkecil ukuran, format, & **Batch Rename Sesuai Excel (Baru)**.</div>', unsafe_allow_html=True)
        if st.button("Buka Kompres Foto", key="dash_foto"):
            navigate_to("Kompres Foto")

    # Kolom 2: PDF Tools
    with cols1[1]:
        st.markdown('<div class="feature-card"><b>📄 PDF Tools</b><br>Gabung, pisah, encrypt, Reorder & **Batch Rename Sesuai Excel (Baru)**.</div>', unsafe_allow_html=True)
        if st.button("Buka PDF Tools", key="dash_pdf"):
            navigate_to("PDF Tools")

    # Kolom 3: MCU Tools
    with cols1[2]:
        st.markdown('<div class="feature-card"><b>🏥 MCU Tools</b><br>Proses Excel + PDF untuk hasil MCU / Analisis Data.</div>', unsafe_allow_html=True)
        if st.button("Buka MCU Tools", key="dash_mcu"):
            navigate_to("MCU Tools")
            
    st.markdown("### Fitur Lainnya")
    cols2 = st.columns(3)
    
    # Kolom 1: File Tools
    with cols2[0]:
        st.markdown('<div class="feature-card"><b>🗂️ File Tools</b><br>Zip/unzip, konversi, **Batch Rename** (duplikasi fitur).</div>', unsafe_allow_html=True)
        if st.button("Buka File Tools", key="dash_file"):
            navigate_to("File Tools")

    # Kolom 2: Tentang
    with cols2[1]:
        st.markdown('<div class="feature-card"><b>ℹ️ Tentang Aplikasi</b><br>Informasi dan kebutuhan library.</div>', unsafe_allow_html=True)
        if st.button("Lihat Tentang", key="dash_about"):
            navigate_to("Tentang")

    with cols2[2]:
        st.markdown('<div class="feature-card" style="visibility:hidden; height: 100%;">.</div>', unsafe_allow_html=True)
        

    st.markdown("---")
    st.info("Semua proses berlangsung lokal di perangkat server tempat Streamlit dijalankan.")


# -------------- 5.2 Kompres Foto / Image Tools --------------
if menu == "Kompres Foto":
    add_back_to_dashboard_button() 
    st.subheader("📸 Kompres & Kelola Foto/Gambar")
    
    img_tool = st.selectbox("Pilih Fitur Gambar", [
        "Kompres Foto (Batch)", 
        "🔢 Batch Rename Gambar (Sequential)", 
        "📄 Batch Rename Gambar Sesuai Excel (Baru)" 
        ], key="img_tool_select")

    st.markdown("---")

    # --- Logika Kompres Foto ---
    if img_tool == "Kompres Foto (Batch)":
        st.markdown("#### 🔄 Kompres Foto/Gambar Massal")
        uploaded = st.file_uploader("Unggah gambar (jpg/png) — bisa banyak", type=["jpg","jpeg","png"], accept_multiple_files=True)
        quality = st.slider("Kualitas JPEG (%)", 10, 95, 75)
        max_side = st.number_input("Max Side (px) [untuk Resize]", min_value=100, max_value=4000, value=1200)
        
        if uploaded and st.button("Proses Kompresi & ZIP", key="process_compress_img"):
            # ... (Logika kompresi, sama seperti sebelumnya) ...
            out_map = {}
            total = len(uploaded)
            prog = st.progress(0)
            with st.spinner(f"Mengompres {total} file..."):
                for i, f in enumerate(uploaded):
                    try:
                        raw = f.read()
                        im = Image.open(io.BytesIO(raw))
                        im.thumbnail((max_side, max_side))
                        buf = io.BytesIO()
                        im.convert("RGB").save(buf, format="JPEG", quality=quality, optimize=True)
                        out_map[f"compressed_{f.name}"] = buf.getvalue()
                    except Exception as e:
                        st.warning(f"Gagal: {f.name} — {e}")
                    prog.progress(int((i+1)/total*100))
            if out_map:
                zipb = make_zip_from_map(out_map)
                st.success(f"✅ {len(out_map)} file berhasil dikompres dan di-ZIP.")
                st.download_button("Unduh Hasil (ZIP)", zipb, file_name="foto_kompres.zip", mime="application/zip")
            else:
                st.warning("Tidak ada file berhasil dikompres.")

    # --- Logika Batch Rename Sequential ---
    elif img_tool == "🔢 Batch Rename Gambar (Sequential)":
        st.markdown("#### 🔢 Ganti Nama Gambar Massal (Urutan Angka)")
        # ... (Logika Batch Rename Sequential, sama seperti sebelumnya) ...
        uploaded_files = st.file_uploader("Unggah file Gambar (JPG, PNG, dll.):", type=["jpg", "jpeg", "png", "webp"], accept_multiple_files=True, key="batch_rename_uploader")
        
        if uploaded_files:
            col1, col2 = st.columns(2)
            new_prefix = col1.text_input("Prefix Nama File Baru:", value="KAY_File", help="Contoh: KAY_File_001.jpg", key="prefix_img_seq")
            new_format = col2.selectbox("Format Output Baru:", ["Sama seperti Asli", "JPG", "PNG", "WEBP"], index=0, key="format_img_seq")
            
            if st.button("Proses Ganti Nama & ZIP", key="process_batch_rename_seq_img"):
                if not new_prefix: st.error("Prefix nama file tidak boleh kosong.")
                else:
                    output_zip = io.BytesIO()
                    try:
                        with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for i, file in enumerate(uploaded_files, 1):
                                # ... (Logika penamaan dan konversi) ...
                                _, original_ext = os.path.splitext(file.name)
                                img = Image.open(file)
                                img_io = io.BytesIO()
                                if new_format == "Sama seperti Asli":
                                    output_format_pil = img.format if img.format else 'JPEG'
                                    output_ext = original_ext
                                else:
                                    output_ext = "." + new_format.lower()
                                    output_format_pil = new_format.upper()
                                new_filename = f"{new_prefix}_{i:03d}{output_ext}"
                                if output_format_pil in ('JPEG', 'JPG'):
                                    img.convert("RGB").save(img_io, format='JPEG', quality=95) 
                                elif output_format_pil == 'PNG':
                                    img.save(img_io, format='PNG')
                                elif output_format_pil == 'WEBP':
                                    img.save(img_io, format='WEBP')
                                else:
                                    img.save(img_io, format=output_format_pil) 
                                img_io.seek(0)
                                zf.writestr(new_filename, img_io.read())
                        st.success(f"✅ Berhasil memproses {len(uploaded_files)} file.")
                        st.download_button("Unduh File ZIP Hasil Batch", data=output_zip.getvalue(), file_name="hasil_batch_gambar.zip", mime="application/zip")
                    except Exception as e: st.error(f"Gagal memproses file: {e}"); traceback.print_exc()

    # --- Logika Batch Rename Sesuai Excel (Paling Penting) ---
    elif img_tool == "📄 Batch Rename Gambar Sesuai Excel (Baru)":
        st.markdown("#### 📄 Ganti Nama Gambar Berdasarkan Daftar Excel")
        st.info("💡 Wajib: Excel/CSV harus punya kolom **`nama_lama`** dan **`nama_baru`** (termasuk ekstensi, misal: `.jpg`).")
        
        col_ex, col_files = st.columns(2)
        excel_up = col_ex.file_uploader("1️⃣ Unggah Excel/CSV Daftar Nama:", type=["xlsx", "csv"], key="rename_img_excel_up")
        files = col_files.file_uploader("2️⃣ Unggah Gambar (JPG/PNG, multiple):", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="rename_img_files_up")
        
        if excel_up and files and st.button("3️⃣ Proses Ganti Nama Gambar (ZIP)", key="process_img_rename_excel"):
            try:
                with st.spinner("Memproses penggantian nama..."):
                    # 1. Baca Excel
                    if excel_up.name.lower().endswith(".csv"):
                        df = pd.read_csv(io.BytesIO(excel_up.read()))
                    else:
                        df = pd.read_excel(io.BytesIO(excel_up.read()))
                    
                    # 2. Validasi Kolom
                    required_cols = ['nama_lama', 'nama_baru']
                    if not all(col in df.columns for col in required_cols):
                        st.error(f"❌ Error: Excel/CSV wajib memiliki kolom: {', '.join(required_cols)}")
                    else:
                        # 3. Map File dan Proses Rename
                        file_map = {f.name: f.read() for f in files}
                        out_map = {}
                        not_found = []

                        for _, row in df.iterrows():
                            old_name = str(row['nama_lama']).strip()
                            new_name = str(row['nama_baru']).strip()
                            
                            if old_name in file_map:
                                if not os.path.splitext(new_name)[1]:
                                    # Tambahkan ekstensi dari nama lama jika nama baru tidak ada ekstensinya
                                    _, old_ext = os.path.splitext(old_name)
                                    new_name = new_name + old_ext
                                    
                                out_map[new_name] = file_map[old_name]
                            else:
                                not_found.append(old_name)

                        # 4. Buat ZIP
                        if out_map:
                            zipb = make_zip_from_map(out_map)
                            st.success(f"✅ Selesai! {len(out_map)} file berhasil diganti namanya dan dikemas.")
                            st.download_button("⬇️ Unduh Hasil (ZIP)", zipb, file_name="gambar_renamed_by_excel.zip", mime="application/zip")
                        else:
                            st.warning("⚠️ Tidak ada file yang cocok ditemukan atau diproses.")
                        
                        if not_found:
                            st.info(f"ℹ️ {len(not_found)} 'nama_lama' di Excel tidak ditemukan di file yang diunggah. (Cek nama dan ekstensi)")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan pemrosesan: {e}"); traceback.print_exc()

# -------------- 5.3 PDF Tools --------------
if menu == "PDF Tools":
    add_back_to_dashboard_button() 
    st.subheader("📄 PDF Tools - Manajemen Dokumen")

    pdf_options = [
        "--- Pilih Tools ---",
        "📂 Gabung PDF",
        "✂️ Pisah PDF", 
        "🔀 Reorder/Hapus Halaman PDF", 
        "📄 Batch Rename PDF Sesuai Excel (Baru)", # FITUR BARU: Batch Rename PDF by Excel
        "📑 Batch Rename PDF (Sequential)", 
        "➡️ Image -> PDF",
        "⬅️ PDF -> Image", 
        "📝 Ekstraksi Teks/Tabel",
        "🔒 Proteksi PDF (Encrypt/Decrypt)",
        "🛠️ Utility PDF (Rotate, Watermark, Kompres)",
    ]
    
    tool_select = st.selectbox("Pilih Fitur PDF:", pdf_options)

    # Simple mapping for internal logic (cleaned up)
    if tool_select.endswith("Excel (Baru)"): tool = "Batch Rename PDF Excel"
    elif tool_select.endswith("(Sequential)"): tool = "Batch Rename PDF Seq"
    elif tool_select.startswith("📂"): tool = "Gabung PDF"
    elif tool_select.startswith("✂️"): tool = "Pisah PDF"
    elif tool_select.startswith("🔀"): tool = "Reorder PDF" 
    elif tool_select.startswith("➡️"): tool = "Image -> PDF"
    elif tool_select.startswith("⬅️"): tool = "PDF -> Image"
    elif tool_select.startswith("📝"): tool = st.selectbox("Pilih mode ekstraksi", ["Extract Text", "Extract Tables -> Excel"], key="pdf_extract_mode")
    elif tool_select.startswith("🔒"): tool = st.selectbox("Pilih mode proteksi", ["Encrypt PDF", "Decrypt PDF", "Batch Lock (Excel)"], key="pdf_protect_mode")
    elif tool_select.startswith("🛠️"): tool = st.selectbox("Pilih mode utilitas", ["Hapus Halaman", "Rotate PDF", "Kompres PDF", "Watermark PDF", "Preview PDF"], key="pdf_util_mode")
    else: tool = None

    st.markdown("---")

    # --- Logika Batch Rename Sesuai Excel (Paling Penting) ---
    if tool == "Batch Rename PDF Excel":
        st.markdown("#### 📄 Ganti Nama File PDF Berdasarkan Daftar Excel")
        st.info("💡 Wajib: Excel/CSV harus punya kolom **`nama_lama`** dan **`nama_baru`** (termasuk ekstensi `.pdf`).")

        col_ex, col_files = st.columns(2)
        excel_up = col_ex.file_uploader("1️⃣ Unggah Excel/CSV Daftar Nama:", type=["xlsx", "csv"], key="rename_pdf_excel_up")
        files = col_files.file_uploader("2️⃣ Unggah File PDF (multiple):", type=["pdf"], accept_multiple_files=True, key="rename_pdf_files_up")
        
        if excel_up and files and st.button("3️⃣ Proses Ganti Nama PDF (ZIP)", key="process_pdf_rename_excel"):
            try:
                with st.spinner("Memproses penggantian nama..."):
                    if excel_up.name.lower().endswith(".csv"):
                        df = pd.read_csv(io.BytesIO(excel_up.read()))
                    else:
                        df = pd.read_excel(io.BytesIO(excel_up.read()))
                    
                    required_cols = ['nama_lama', 'nama_baru']
                    if not all(col in df.columns for col in required_cols):
                        st.error(f"❌ Error: Excel/CSV wajib memiliki kolom: {', '.join(required_cols)}")
                    else:
                        file_map = {f.name: f.read() for f in files}
                        out_map = {}
                        not_found = []
                        
                        for _, row in df.iterrows():
                            old_name = str(row['nama_lama']).strip()
                            new_name = str(row['nama_baru']).strip()
                            
                            if old_name in file_map:
                                if not new_name.lower().endswith('.pdf'):
                                    new_name += '.pdf'
                                out_map[new_name] = file_map[old_name]
                            else:
                                not_found.append(old_name)

                        if out_map:
                            zipb = make_zip_from_map(out_map)
                            st.success(f"✅ Selesai! {len(out_map)} file berhasil diganti namanya dan dikemas.")
                            st.download_button("⬇️ Unduh Hasil (ZIP)", zipb, file_name="pdf_renamed_by_excel.zip", mime="application/zip")
                        else:
                            st.warning("⚠️ Tidak ada file yang cocok ditemukan atau diproses.")
                        
                        if not_found:
                            st.info(f"ℹ️ {len(not_found)} 'nama_lama' di Excel tidak ditemukan di file yang diunggah. (Cek nama dan ekstensi)")
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan pemrosesan: {e}"); traceback.print_exc()

    # --- Logika Batch Rename Sequential PDF ---
    if tool == "Batch Rename PDF Seq":
        st.markdown("#### 📑 Ganti Nama File PDF Massal (Urutan Angka)")
        uploaded_files = st.file_uploader("Unggah file PDF (multiple):", type=["pdf"], accept_multiple_files=True, key="batch_rename_pdf_uploader_seq")
        
        if uploaded_files:
            col1, col2 = st.columns(2)
            new_prefix = col1.text_input("Prefix Nama File Baru:", value="Hasil_PDF", help="Contoh: Hasil_PDF_001.pdf", key="prefix_pdf_seq")
            start_num = col2.number_input("Mulai dari Angka (Counter):", min_value=1, value=1, step=1, key="start_num_pdf_seq")

            if st.button("Proses Ganti Nama & ZIP", key="process_batch_rename_pdf_seq"):
                if not new_prefix: st.error("Prefix nama file tidak boleh kosong.")
                else:
                    output_zip = io.BytesIO()
                    try:
                        with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for i, file in enumerate(uploaded_files, start_num):
                                new_filename = f"{new_prefix}_{i:03d}.pdf"
                                zf.writestr(new_filename, file.read())
                        st.success(f"✅ Berhasil mengganti nama {len(uploaded_files)} file.")
                        st.download_button("⬇️ Unduh File ZIP Hasil Rename", data=output_zip.getvalue(), file_name="pdf_renamed.zip", mime="application/zip")
                    except Exception as e: st.error(f"❌ Gagal memproses file: {e}"); traceback.print_exc()

    # --- Logika Reorder/Hapus Halaman PDF ---
    if tool == "Reorder PDF":
        st.markdown("#### 🔀 Reorder atau Hapus Halaman PDF")
        st.markdown("Tentukan urutan halaman baru, dipisahkan koma (contoh: `2, 1, 3`). Menghilangkan angka berarti menghapus halaman tersebut.")

        f = st.file_uploader("Unggah 1 file PDF:", type="pdf", key="reorder_pdf_uploader")
        
        if f:
            try:
                raw = f.read()
                reader = PdfReader(io.BytesIO(raw))
                num_pages = len(reader.pages)
                st.info(f"PDF berhasil dimuat. Total halaman: **{num_pages}**.")
                
                default_order = ", ".join(map(str, range(1, num_pages + 1)))

                new_order_str = st.text_input(
                    f"Masukkan Urutan Halaman Baru (1-{num_pages}):",
                    value=default_order,
                    help="Contoh: '3, 1, 2' (balik urutan 3 halaman). '1, 3, 5' (ambil halaman ganjil)."
                )

                if st.button("Proses Reorder/Hapus", key="process_reorder"):
                    
                    try:
                        input_list = [int(x.strip()) for x in new_order_str.split(',') if x.strip().isdigit()]
                        
                        if any(n < 1 or n > num_pages for n in input_list):
                            st.error(f"❌ Error: Nomor halaman harus antara 1 sampai {num_pages}.")
                            raise ValueError("Invalid page number in input.")

                        new_order_indices = [n - 1 for n in input_list]
                        
                        writer = PdfWriter()
                        for index in new_order_indices:
                            writer.add_page(reader.pages[index])
                        
                        pdf_buffer = io.BytesIO(); writer.write(pdf_buffer); pdf_buffer.seek(0)

                        st.download_button("✅ Unduh Hasil PDF (Reordered)", data=pdf_buffer, file_name="pdf_reordered.pdf", mime="application/pdf")
                        st.success(f"Pemrosesan selesai. Total halaman baru: {len(new_order_indices)}.")

                    except ValueError:
                        st.error("❌ Error: Format urutan halaman tidak valid. Pastikan hanya angka dan koma.")
                    except Exception as e:
                        st.error(f"❌ Terjadi kesalahan pemrosesan PDF: {e}")

            except Exception as e:
                st.error(f"❌ Terjadi kesalahan saat memproses PDF: {e}"); st.info("Pastikan file yang diunggah adalah PDF yang valid.")

    # --- Logika Fitur PDF Lainnya (Gabung/Pisah/Encrypt, dll.) ---
    # Logika Gabung PDF (Sama seperti sebelumnya)
    if tool == "Gabung PDF":
        st.markdown("#### 🔗 Gabungkan Beberapa File PDF")
        files = st.file_uploader("Unggah file PDF (multiple):", type="pdf", accept_multiple_files=True)
        if files and st.button("Proses Gabung"):
            try:
                with st.spinner("Menggabungkan..."):
                    writer = PdfWriter()
                    for f in files:
                        reader = PdfReader(io.BytesIO(f.read()));
                        for p in reader.pages: writer.add_page(p)
                    out = io.BytesIO(); writer.write(out); out.seek(0)
                st.download_button("⬇️ Download merged.pdf", out.getvalue(), file_name="merged.pdf", mime="application/pdf")
                st.success("✅ Gabung Selesai")
            except Exception: st.error(traceback.format_exc())

    # Logika Pisah PDF (Sama seperti sebelumnya)
    if tool == "Pisah PDF":
        st.markdown("#### ✂️ Pisahkan PDF Menjadi Per Halaman")
        f = st.file_uploader("Unggah 1 file PDF:", type="pdf")
        if f and st.button("Proses Pisah (ZIP)"):
            # ... (Logika pisah, sama seperti sebelumnya) ...
            try:
                with st.spinner("Memisahkan..."):
                    reader = PdfReader(io.BytesIO(f.read()))
                    out_map = {}
                    for i, p in enumerate(reader.pages):
                        w = PdfWriter(); w.add_page(p)
                        buf = io.BytesIO(); w.write(buf); buf.seek(0)
                        out_map[f"page_{i+1}.pdf"] = buf.getvalue()
                    zipb = make_zip_from_map(out_map)
                st.download_button("⬇️ Download pages.zip", zipb, file_name="pages.zip", mime="application/zip")
                st.success("✅ Pemisahan Selesai")
            except Exception: st.error(traceback.format_exc())

    # Logika Enkripsi/Dekripsi/Batch Lock (Diambil dari sub-menu "🔒 Proteksi")
    if tool in ["Encrypt PDF", "Decrypt PDF", "Batch Lock (Excel)"]:
        st.markdown(f"#### {tool_select}")
        if tool == "Encrypt PDF":
            f = st.file_uploader("Unggah PDF", type="pdf")
            pw = st.text_input("Password", type="password")
            if f and pw and st.button("Proses Enkripsi"):
                try:
                    with st.spinner("Mengunci PDF..."):
                        reader = PdfReader(io.BytesIO(f.read())); writer = PdfWriter();
                        for p in reader.pages: writer.add_page(p)
                        try_encrypt(writer, pw); buf = io.BytesIO(); writer.write(buf); buf.seek(0)
                    st.download_button("⬇️ Download encrypted.pdf", buf.getvalue(), file_name="encrypted.pdf", mime="application/pdf")
                except Exception: st.error(traceback.format_exc())

        if tool == "Decrypt PDF":
            f = st.file_uploader("Unggah encrypted PDF", type="pdf")
            pw = st.text_input("Password untuk dekripsi", type="password")
            if f and pw and st.button("Proses Dekripsi"):
                try:
                    with st.spinner("Membuka PDF..."):
                        reader = PdfReader(io.BytesIO(f.read()))
                        if getattr(reader, "is_encrypted", False): reader.decrypt(pw)
                        writer = PdfWriter();
                        for p in reader.pages: writer.add_page(p)
                        buf = io.BytesIO(); writer.write(buf); buf.seek(0)
                    st.download_button("⬇️ Download decrypted.pdf", buf.getvalue(), file_name="decrypted.pdf", mime="application/pdf")
                except Exception: st.error(traceback.format_exc())
        
        if tool == "Batch Lock (Excel)":
            st.info("💡 Wajib: Excel/CSV harus punya kolom `filename` dan `password`.")
            excel_file = st.file_uploader("1️⃣ Unggah Excel (filename,password):", type=["xlsx","csv"])
            pdfs = st.file_uploader("2️⃣ Unggah PDFs (multiple):", type="pdf", accept_multiple_files=True)
            if excel_file and pdfs and st.button("3️⃣ Proses Batch Lock (ZIP)"):
                # ... (Logika Batch Lock, sama seperti sebelumnya) ...
                try:
                    with st.spinner("Batch locking PDFs..."):
                        if excel_file.name.lower().endswith(".csv"): df = pd.read_csv(io.BytesIO(excel_file.read()))
                        else: df = pd.read_excel(io.BytesIO(excel_file.read()))
                        pdf_map = {p.name: p.read() for p in pdfs}; out_map = {}; not_found = []; total = len(df); prog = st.progress(0)
                        
                        cols = [c.lower() for c in df.columns]
                        if 'filename' not in cols or 'password' not in cols:
                            st.error("❌ Error: Excel harus mengandung kolom 'filename' dan 'password'.")
                        else:
                            fn_col = df.columns[cols.index('filename')]; pwd_col = df.columns[cols.index('password')]
                            for idx, (_, row) in enumerate(df.iterrows()):
                                target = str(row[fn_col]).strip(); pwd = str(row[pwd_col]).strip()
                                if target and pwd and target in pdf_map:
                                    reader = PdfReader(io.BytesIO(pdf_map[target])); writer = PdfWriter();
                                    for p in reader.pages: writer.add_page(p)
                                    try_encrypt(writer, pwd); b = io.BytesIO(); writer.write(b); out_map[f"locked_{target}"] = b.getvalue()
                                else: not_found.append(target)
                                prog.progress(int((idx+1)/total*100))

                    if out_map:
                        st.download_button("⬇️ Download locked_pdfs.zip", make_zip_from_map(out_map), file_name="locked_pdfs.zip", mime="application/zip")
                        st.success(f"✅ {len(out_map)} file berhasil dikunci.")
                    if not_found: st.warning(f"⚠️ {len(not_found)} files tidak ditemukan. Sampel: {not_found[:5]}")
                except Exception: st.error(traceback.format_exc())

    # Logika Ekstraksi Teks/Tabel (Diambil dari sub-menu "📝 Ekstraksi")
    if tool in ["Extract Text", "Extract Tables -> Excel"]:
        st.markdown(f"#### {tool_select}")
        # ... (Logika Ekstraksi, sama seperti sebelumnya) ...
        f = st.file_uploader("Unggah PDF:", type="pdf")
        if f and st.button(f"Proses {tool}"):
            try:
                if tool == "Extract Text":
                    with st.spinner("Mengekstrak teks..."):
                        # Logika extract text...
                        text_blocks = []; raw = f.read()
                        reader = PdfReader(io.BytesIO(raw))
                        for i, p in enumerate(reader.pages): text_blocks.append(f"--- Page {i+1} ---\n" + (p.extract_text() or ""))
                        full = "\n".join(text_blocks)
                    st.text_area("Extracted text (preview)", full[:10000], height=300)
                    st.download_button("⬇️ Download .txt", full, file_name="extracted_text.txt", mime="text/plain")
                    st.success("✅ Ekstraksi Teks Selesai")

                elif tool == "Extract Tables -> Excel":
                    if pdfplumber is None: st.error("❌ Error: `pdfplumber` diperlukan untuk ekstraksi tabel.")
                    else:
                        with st.spinner("Mengekstrak tabel..."):
                            tables = [];
                            with pdfplumber.open(io.BytesIO(f.read())) as doc:
                                for page in doc.pages:
                                    for tbl in page.extract_tables():
                                        if tbl and len(tbl) > 1:
                                            df = pd.DataFrame(tbl[1:], columns=tbl[0]) 
                                            tables.append(df)
                        if tables:
                            df_all = pd.concat(tables, ignore_index=True)
                            st.dataframe(df_all.head()); excel_bytes = df_to_excel_bytes(df_all)
                            st.download_button("⬇️ Download Excel", data=excel_bytes, file_name="extracted_tables.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                            st.success(f"✅ Ekstraksi Tabel Selesai. Total {len(tables)} tabel ditemukan.")
                        else:
                            st.warning("⚠️ Tidak ada tabel yang ditemukan.")

            except Exception: st.error(traceback.format_exc())

    # Logika Utilitas PDF (Diambil dari sub-menu "🛠️ Utility")
    if tool in ["Hapus Halaman", "Rotate PDF", "Kompres PDF", "Watermark PDF", "Preview PDF"]:
        st.markdown(f"#### 🛠️ {tool_select}")
        # ... (Logika Utilitas, sama seperti sebelumnya) ...
        if tool == "Hapus Halaman":
            f = st.file_uploader("Unggah PDF", type="pdf")
            page_no = st.number_input("Nomor Halaman yang Dihapus (1-based)", min_value=1, value=1)
            if f and st.button("Proses Hapus Halaman"):
                try:
                    with st.spinner("Menghapus..."):
                        reader = PdfReader(io.BytesIO(f.read())); writer = PdfWriter()
                        for i, p in enumerate(reader.pages):
                            if i+1 != page_no: writer.add_page(p)
                        buf = io.BytesIO(); writer.write(buf); buf.seek(0)
                    st.download_button("⬇️ Download result", buf.getvalue(), file_name="removed_page.pdf", mime="application/pdf")
                    st.success("✅ Penghapusan Selesai")
                except Exception: st.error(traceback.format_exc())
        
        # ... (Logika Rotate, Kompres, Watermark, Preview, sama seperti sebelumnya) ...
        if tool == "Rotate PDF":
            f = st.file_uploader("Unggah PDF", type="pdf")
            angle = st.selectbox("Sudut Rotasi", [90, 180, 270])
            if f and st.button("Proses Rotasi"):
                try:
                    with st.spinner("Memutar..."):
                        reader = PdfReader(io.BytesIO(f.read())); writer = PdfWriter()
                        for p in reader.pages:
                            try: p.rotate(angle)
                            except: pass
                            writer.add_page(p)
                        buf = io.BytesIO(); writer.write(buf); buf.seek(0)
                    st.download_button("⬇️ Download rotated.pdf", buf.getvalue(), file_name="rotated.pdf", mime="application/pdf")
                    st.success("✅ Rotasi Selesai")
                except Exception: st.error(traceback.format_exc())

        if tool == "Kompres PDF":
            f = st.file_uploader("Unggah PDF", type="pdf")
            if f and st.button("Proses Kompresi (Rewrite)"):
                try:
                    with st.spinner("Mengompres (rewrite)..."):
                        reader = PdfReader(io.BytesIO(f.read())); writer = PdfWriter()
                        for p in reader.pages: writer.add_page(p)
                        buf = io.BytesIO(); writer.write(buf); buf.seek(0)
                    st.download_button("⬇️ Download compressed.pdf", buf.getvalue(), file_name="compressed.pdf", mime="application/pdf")
                    st.success("✅ Kompresi Selesai")
                except Exception: st.error(traceback.format_exc())

        if tool == "Watermark PDF":
            base = st.file_uploader("1️⃣ Base PDF", type="pdf")
            watermark = st.file_uploader("2️⃣ Watermark PDF (1 Halaman)", type="pdf")
            if base and watermark and st.button("Proses Watermark"):
                try:
                    with st.spinner("Menerapkan watermark..."):
                        rb = PdfReader(io.BytesIO(base.read())); rm = PdfReader(io.BytesIO(watermark.read())); wm = rm.pages[0]; writer = PdfWriter()
                        for p in rb.pages:
                            try: p.merge_page(wm)
                            except: pass
                            writer.add_page(p)
                        buf = io.BytesIO(); writer.write(buf); buf.seek(0)
                    st.download_button("⬇️ Download watermarked.pdf", buf.getvalue(), file_name="watermarked.pdf", mime="application/pdf")
                    st.success("✅ Watermark Selesai")
                except Exception: st.error(traceback.format_exc())

        if tool == "Preview PDF":
            f = st.file_uploader("Unggah PDF", type="pdf")
            mode = st.radio("Mode Preview", ["First page (fast)", "All pages (slow)"])
            if f and st.button("Tampilkan Preview"):
                # ... (Logika preview, sama seperti sebelumnya) ...
                try:
                    with st.spinner("Preparing preview..."):
                        pdf_bytes = f.read()
                        if PDF2IMAGE_AVAILABLE:
                            # Logika konversi ke gambar
                            images = None
                            if mode.startswith("First"):
                                if convert_from_bytes is not None: images = convert_from_bytes(pdf_bytes, first_page=1, last_page=1)
                                else: # Fallback using temp file for convert_from_path
                                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp: tmp.write(pdf_bytes); tmp_path = tmp.name
                                    images = convert_from_path(tmp_path, first_page=1, last_page=1); os.unlink(tmp_path)
                                buf = io.BytesIO(); images[0].save(buf, format="PNG"); buf.seek(0); st.image(buf.getvalue(), caption="Page 1")
                            else:
                                if convert_from_bytes is not None: images = convert_from_bytes(pdf_bytes)
                                else: # Fallback using temp file
                                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp: tmp.write(pdf_bytes); tmp_path = tmp.name
                                    images = convert_from_path(tmp_path); os.unlink(tmp_path)
                                for i, img in enumerate(images):
                                    buf = io.BytesIO(); img.save(buf, format="PNG"); st.image(buf.getvalue(), caption=f"Page {i+1}")
                        else:
                            st.warning("⚠️ `pdf2image` tidak terinstal, menampilkan teks saja.")
                            reader = PdfReader(io.BytesIO(pdf_bytes))
                            if mode.startswith("First"): st.text(reader.pages[0].extract_text() or "[no text]")
                            else:
                                for i, p in enumerate(reader.pages): st.write(f"--- Page {i+1} ---"); st.text(p.extract_text() or "[no text]")
                    st.success("✅ Preview Selesai")
                except Exception: st.error(traceback.format_exc())

    # Logika PDF -> Image / Image -> PDF
    if tool == "PDF -> Image":
        st.markdown("#### ⬅️ Konversi PDF ke Gambar (Batch)")
        st.info("⚠️ Membutuhkan `pdf2image` dan `poppler` terinstal di server.")
        f = st.file_uploader("Unggah PDF", type="pdf")
        dpi = st.slider("DPI (Kualitas Gambar)", 100, 300, 150)
        fmt = st.radio("Format Gambar Output", ["PNG", "JPEG"])
        if f and st.button("Proses Konversi ke Gambar"):
            # ... (Logika PDF -> Image, sama seperti sebelumnya) ...
            try:
                if not PDF2IMAGE_AVAILABLE: st.error("❌ Error: `pdf2image` tidak terinstal atau `poppler` hilang.")
                else:
                    with st.spinner("Converting..."):
                        pdf_bytes = f.read(); images = None
                        if convert_from_bytes is not None: images = convert_from_bytes(pdf_bytes, dpi=dpi)
                        else: # Fallback using temp file
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp: tmp.write(pdf_bytes); tmp_path = tmp.name
                            images = convert_from_path(tmp_path, dpi=dpi); os.unlink(tmp_path)
                        out_map = {};
                        for i, img in enumerate(images):
                            b = io.BytesIO(); img.save(b, format=fmt); out_map[f"page_{i+1}.{fmt.lower()}"] = b.getvalue()
                        zipb = make_zip_from_map(out_map)
                    st.download_button("⬇️ Download images.zip", zipb, file_name="pdf_images.zip", mime="application/zip")
                    st.success("✅ Konversi Selesai")
            except Exception: st.error(traceback.format_exc())
    
    if tool == "Image -> PDF":
        st.markdown("#### ➡️ Konversi Gambar ke PDF (Gabung)")
        imgs = st.file_uploader("Unggah Gambar (multiple)", type=["jpg","png","jpeg"], accept_multiple_files=True)
        if imgs and st.button("Proses Gambar -> PDF"):
            # ... (Logika Image -> PDF, sama seperti sebelumnya) ...
            try:
                with st.spinner("Membuat PDF dari gambar..."):
                    pil = [Image.open(io.BytesIO(i.read())).convert("RGB") for i in imgs]
                    buf = io.BytesIO()
                    if len(pil) == 1: pil[0].save(buf, format="PDF")
                    else: pil[0].save(buf, save_all=True, append_images=pil[1:], format="PDF")
                    buf.seek(0)
                st.download_button("⬇️ Download images_as_pdf.pdf", buf.getvalue(), file_name="images_as_pdf.pdf", mime="application/pdf")
                st.success("✅ Konversi Selesai")
            except Exception: st.error(traceback.format_exc())

# -------------- 5.4 MCU Tools --------------
if menu == "MCU Tools":
    add_back_to_dashboard_button() 
    st.subheader("🏥 MCU Tools - Organise & Analyze Data")

    mcu_mode = st.selectbox("Pilih Mode MCU", [
        "Organise Files by Excel (Struktur Folder)", 
        "📊 Analisis Data MCU Massal (Dashboard)",
        ], key="mcu_mode_select")
    
    st.markdown("---")
    
    # --- Logika Analisis Data MCU Massal ---
    if mcu_mode == "📊 Analisis Data MCU Massal (Dashboard)":
        st.markdown("#### 📊 Dashboard Analisis Hasil MCU Massal")
        # ... (Logika Analisis, sama seperti sebelumnya) ...
        uploaded_file = st.file_uploader("Unggah file Data MCU (Excel/CSV):", type=["xlsx", "csv"], key="mcu_data_uploader")

        if uploaded_file:
            try:
                if uploaded_file.name.lower().endswith('.csv'): df = pd.read_csv(uploaded_file)
                else: df = pd.read_excel(uploaded_file)

                st.success(f"File **{uploaded_file.name}** berhasil dimuat. Total Baris: {len(df)}")
                df.columns = df.columns.str.replace('[^A-Za-z0-9_]+', '', regex=True).str.lower()
                st.dataframe(df.head(), use_container_width=True)

                st.markdown("#### 📈 Hasil Analisis Agregat")
                status_cols = [col for col in df.columns if 'status' in col or 'fit' in col]
                
                if status_cols:
                    status_col = status_cols[0]
                    st.write(f"##### 1. Distribusi Status Kesehatan (Kolom: `{status_col}`)")
                    df[status_col] = df[status_col].fillna("TIDAK DIKETAHUI") 
                    status_counts = df[status_col].value_counts().reset_index()
                    status_counts.columns = ['Status', 'Jumlah']
                    
                    if len(status_counts) > 0: st.bar_chart(status_counts.set_index('Status'), color="#4CAF50")
                    else: st.info("Tidak ada data unik yang valid dalam kolom status.")
                else: st.warning("⚠️ Kolom status ('status' atau 'fit') tidak ditemukan untuk Analisis Cepat.")
            except Exception as e: st.error(f"❌ Gagal memuat atau memproses file: {e}"); traceback.print_exc()

    # --- Logika Organise Files by Excel ---
    if mcu_mode == "Organise Files by Excel (Struktur Folder)":
        st.markdown("#### 📂 Atur File PDF Sesuai Struktur Excel")
        st.info("💡 Excel bisa berisi: (A) `No_MCU`, `Nama`, `Departemen`, `JABATAN` atau (B) `filename`, `target_folder`.")
        excel_up = st.file_uploader("1️⃣ Unggah Excel Daftar Organisasi:", type=["xlsx","csv"], key="mcu_organize_excel")
        pdfs = st.file_uploader("2️⃣ Unggah File PDF MCU (multiple):", type="pdf", accept_multiple_files=True, key="mcu_organize_pdf")
        
        if excel_up and pdfs and st.button("3️⃣ Proses Organisasi (ZIP)"):
            # ... (Logika Organise, sama seperti sebelumnya) ...
            try:
                with st.spinner("Memproses Organisasi..."):
                    if excel_up.name.lower().endswith(".csv"): df = pd.read_csv(io.BytesIO(excel_up.read()))
                    else: df = pd.read_excel(io.BytesIO(excel_up.read()))
                
                    pdf_map = {p.name: p.read() for p in pdfs}; out_map = {}; not_found = []
                    
                    if all(c in df.columns for c in ["No_MCU","Nama","Departemen","JABATAN"]):
                        # Mode A: Organisasi berdasarkan kolom MCU
                        for _, r in df.iterrows():
                            no = str(r["No_MCU"]).strip()
                            dept = str(r["Departemen"]) if not pd.isna(r["Departemen"]) else "Unknown"
                            jab = str(r["JABATAN"]) if not pd.isna(r["JABATAN"]) else "Unknown"
                            matches = [k for k in pdf_map.keys() if k.startswith(no)]
                            if matches: out_map[f"{dept}/{jab}/{matches[0]}"] = pdf_map[matches[0]]
                            else: not_found.append(no)
                    elif "filename" in df.columns and "target_folder" in df.columns:
                        # Mode B: Organisasi berdasarkan kolom filename dan target_folder
                        for _, r in df.iterrows():
                            fn = str(r["filename"]).strip(); tgt = str(r["target_folder"]).strip()
                            if fn in pdf_map: out_map[f"{tgt}/{fn}"] = pdf_map[fn]
                            else: not_found.append(fn)
                    else:
                        st.error("❌ Error: Kolom Excel tidak sesuai. Gunakan set A atau B.")
                        raise ValueError("Invalid Excel columns.")
                
                if out_map:
                    st.download_button("⬇️ Download MCU zip", make_zip_from_map(out_map), file_name="mcu_structured.zip", mime="application/zip")
                    st.success(f"✅ {len(out_map)} file berhasil diorganisasi ke dalam ZIP.")
                if not_found: st.warning(f"⚠️ {len(not_found)} file/nomor MCU tidak ditemukan. Sampel: {not_found[:5]}")
            except Exception: st.error(traceback.format_exc())

# -------------- 5.5 File Tools --------------
if menu == "File Tools":
    add_back_to_dashboard_button() 
    st.subheader("🗂️ File Tools - Zip / Unzip / Konversi / Rename")
    mode = st.selectbox("Pilih Alat File:", [
        "Zip files (Kompres ke ZIP)", 
        "Unzip file (Ekstrak ZIP)", 
        "Excel -> CSV", 
        "Word -> PDF (Konversi Teks)", 
        "🔢 Batch Rename/Format Gambar (Sequential)", 
        "📄 Batch Rename Gambar Sesuai Excel", 
        "📑 Batch Rename PDF (Sequential)", 
        "📄 Batch Rename PDF Sesuai Excel",
        ], key="file_tool_select")
    
    st.markdown("---")
    
    # --- Duplikasi Logika Batch Rename Sesuai Excel (Untuk Gambar dan PDF) ---
    if mode == "📄 Batch Rename PDF Sesuai Excel":
        # ... (Logika Batch Rename PDF Excel, sama seperti di PDF Tools) ...
        st.markdown("#### 📄 Ganti Nama File PDF Berdasarkan Daftar Excel")
        st.info("💡 Wajib: Excel/CSV harus punya kolom **`nama_lama`** dan **`nama_baru`** (termasuk ekstensi `.pdf`).")
        col_ex, col_files = st.columns(2)
        excel_up = col_ex.file_uploader("1️⃣ Unggah Excel/CSV Daftar Nama:", type=["xlsx", "csv"], key="rename_pdf_excel_up_2")
        files = col_files.file_uploader("2️⃣ Unggah File PDF (multiple):", type=["pdf"], accept_multiple_files=True, key="rename_pdf_files_up_2")
        if excel_up and files and st.button("3️⃣ Proses Ganti Nama PDF (ZIP)", key="process_pdf_rename_excel_2"):
            # (Logika pemrosesan Batch Rename PDF Excel)
            try:
                if excel_up.name.lower().endswith(".csv"): df = pd.read_csv(io.BytesIO(excel_up.read()))
                else: df = pd.read_excel(io.BytesIO(excel_up.read()))
                required_cols = ['nama_lama', 'nama_baru']
                if not all(col in df.columns for col in required_cols): st.error(f"❌ Error: Excel/CSV wajib memiliki kolom: {', '.join(required_cols)}")
                else:
                    file_map = {f.name: f.read() for f in files}; out_map = {}; not_found = []
                    for _, row in df.iterrows():
                        old_name = str(row['nama_lama']).strip(); new_name = str(row['nama_baru']).strip()
                        if old_name in file_map:
                            if not new_name.lower().endswith('.pdf'): new_name += '.pdf'
                            out_map[new_name] = file_map[old_name]
                        else: not_found.append(old_name)
                    if out_map:
                        st.download_button("⬇️ Unduh File ZIP Hasil Rename", data=make_zip_from_map(out_map), file_name="pdf_renamed_2.zip", mime="application/zip")
                        st.success(f"✅ {len(out_map)} file berhasil diganti namanya.")
                    if not_found: st.info(f"ℹ️ {len(not_found)} 'nama_lama' di Excel tidak ditemukan. Sampel: {not_found[:5]}")
            except Exception as e: st.error(f"❌ Terjadi kesalahan pemrosesan: {e}"); traceback.print_exc()

    if mode == "📄 Batch Rename Gambar Sesuai Excel":
        # ... (Logika Batch Rename Gambar Excel, sama seperti di Kompres Foto) ...
        st.markdown("#### 📄 Ganti Nama Gambar Berdasarkan Daftar Excel")
        st.info("💡 Wajib: Excel/CSV harus punya kolom **`nama_lama`** dan **`nama_baru`** (termasuk ekstensi, misal: `.jpg`).")
        col_ex, col_files = st.columns(2)
        excel_up = col_ex.file_uploader("1️⃣ Unggah Excel/CSV Daftar Nama:", type=["xlsx", "csv"], key="rename_img_excel_up_2")
        files = col_files.file_uploader("2️⃣ Unggah Gambar (JPG/PNG/JPEG, multiple):", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="rename_img_files_up_2")
        if excel_up and files and st.button("3️⃣ Proses Ganti Nama Gambar (ZIP)", key="process_img_rename_excel_2"):
            # (Logika pemrosesan Batch Rename Gambar Excel)
            try:
                if excel_up.name.lower().endswith(".csv"): df = pd.read_csv(io.BytesIO(excel_up.read()))
                else: df = pd.read_excel(io.BytesIO(excel_up.read()))
                required_cols = ['nama_lama', 'nama_baru']
                if not all(col in df.columns for col in required_cols): st.error(f"❌ Error: Excel/CSV wajib memiliki kolom: {', '.join(required_cols)}")
                else:
                    file_map = {f.name: f.read() for f in files}; out_map = {}; not_found = []
                    for _, row in df.iterrows():
                        old_name = str(row['nama_lama']).strip(); new_name = str(row['nama_baru']).strip()
                        if old_name in file_map:
                            if not os.path.splitext(new_name)[1]: # Tambah ekstensi dari nama lama
                                _, old_ext = os.path.splitext(old_name); new_name = new_name + old_ext
                            out_map[new_name] = file_map[old_name]
                        else: not_found.append(old_name)
                    if out_map:
                        st.download_button("⬇️ Unduh Hasil (ZIP)", data=make_zip_from_map(out_map), file_name="gambar_renamed_by_excel_2.zip", mime="application/zip")
                        st.success(f"✅ {len(out_map)} file berhasil diganti namanya.")
                    if not_found: st.info(f"ℹ️ {len(not_found)} 'nama_lama' di Excel tidak ditemukan. Sampel: {not_found[:5]}")
            except Exception as e: st.error(f"❌ Terjadi kesalahan pemrosesan: {e}"); traceback.print_exc()

    # --- Duplikasi Logika Batch Rename Sequential ---
    if mode == "📑 Batch Rename PDF (Sequential)":
        st.markdown("#### 📑 Ganti Nama File PDF Massal (Urutan Angka)")
        uploaded_files = st.file_uploader("Unggah file PDF (multiple):", type=["pdf"], accept_multiple_files=True, key="batch_rename_pdf_uploader_file_tool")
        if uploaded_files:
            col1, col2 = st.columns(2)
            new_prefix = col1.text_input("Prefix Nama File Baru:", value="Hasil_PDF", help="Contoh: Hasil_PDF_001.pdf", key="prefix_pdf_file_tool")
            start_num = col2.number_input("Mulai dari Angka:", min_value=1, value=1, step=1, key="start_num_pdf_file_tool")
            if st.button("Proses Ganti Nama (ZIP)", key="process_batch_rename_pdf_file_tool"):
                if not new_prefix: st.error("Prefix nama file tidak boleh kosong.")
                else:
                    output_zip = io.BytesIO()
                    try:
                        with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for i, file in enumerate(uploaded_files, start_num):
                                zf.writestr(f"{new_prefix}_{i:03d}.pdf", file.read())
                        st.success(f"✅ Berhasil mengganti nama {len(uploaded_files)} file."); st.download_button("⬇️ Unduh File ZIP", data=output_zip.getvalue(), file_name="pdf_renamed_file_tool.zip", mime="application/zip")
                    except Exception as e: st.error(f"❌ Gagal memproses file: {e}"); traceback.print_exc()

    if mode == "🔢 Batch Rename/Format Gambar (Sequential)":
        st.markdown("#### 🔢 Ganti Nama Gambar Massal (Urutan Angka)")
        uploaded_files = st.file_uploader("Unggah file Gambar (JPG, PNG, dll.):", type=["jpg", "jpeg", "png", "webp"], accept_multiple_files=True, key="batch_rename_uploader_file_tool")
        if uploaded_files:
            col1, col2 = st.columns(2)
            new_prefix = col1.text_input("Prefix Nama File Baru:", value="KAY_File", help="Contoh: KAY_File_001.jpg", key="prefix_img_file_tool")
            new_format = col2.selectbox("Format Output Baru:", ["Sama seperti Asli", "JPG", "PNG", "WEBP"], index=0, key="format_img_file_tool")
            if st.button("Proses Batch File", key="process_batch_rename_file_tool"):
                if not new_prefix: st.error("Prefix nama file tidak boleh kosong.")
                else:
                    output_zip = io.BytesIO()
                    try:
                        with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for i, file in enumerate(uploaded_files, 1):
                                # (Logika penamaan dan konversi)
                                _, original_ext = os.path.splitext(file.name); img = Image.open(file); img_io = io.BytesIO()
                                output_ext = "." + new_format.lower() if new_format != "Sama seperti Asli" else original_ext
                                output_format_pil = new_format.upper() if new_format != "Sama seperti Asli" else (img.format if img.format else 'JPEG')
                                new_filename = f"{new_prefix}_{i:03d}{output_ext}"
                                # Save logic (same as before)
                                if output_format_pil in ('JPEG', 'JPG'): img.convert("RGB").save(img_io, format='JPEG', quality=95) 
                                elif output_format_pil == 'PNG': img.save(img_io, format='PNG')
                                elif output_format_pil == 'WEBP': img.save(img_io, format='WEBP')
                                else: img.save(img_io, format=output_format_pil) 
                                img_io.seek(0); zf.writestr(new_filename, img_io.read())
                        st.success(f"✅ Berhasil memproses {len(uploaded_files)} file."); st.download_button("⬇️ Unduh File ZIP", data=output_zip.getvalue(), file_name="hasil_batch_gambar_file_tool.zip", mime="application/zip")
                    except Exception as e: st.error(f"❌ Gagal memproses file: {e}"); traceback.print_exc()

    # --- Logika Fitur Zip/Unzip/Konversi ---
    if mode == "Zip files (Kompres ke ZIP)":
        st.markdown("#### 📁 Kompres File ke Format ZIP")
        ups = st.file_uploader("Pilih file yang akan di-ZIP (multiple):", accept_multiple_files=True)
        if ups and st.button("Proses Create ZIP"):
            try:
                # (Logika Zip files, sama seperti sebelumnya)
                with st.spinner("Membuat ZIP..."):
                    out = io.BytesIO();
                    with zipfile.ZipFile(out, "w") as z:
                        for i, f in enumerate(ups): z.writestr(f.name, f.read())
                        out.seek(0)
                st.download_button("⬇️ Download ZIP", out.getvalue(), file_name="files.zip", mime="application/zip")
                st.success("✅ ZIP Selesai")
            except Exception: st.error(traceback.format_exc())

    elif mode == "Unzip file (Ekstrak ZIP)":
        st.markdown("#### 📂 Ekstrak File dari Format ZIP")
        zf = st.file_uploader("Unggah file ZIP:", type="zip")
        if zf and st.button("Proses Ekstrak"):
            try:
                # (Logika Unzip file, sama seperti sebelumnya)
                with st.spinner("Mengekstrak ZIP..."):
                    tmpdir = tempfile.mkdtemp()
                    with zipfile.ZipFile(io.BytesIO(zf.read()), "r") as z: z.extractall(tmpdir)
                    shutil.make_archive(tmpdir, "zip", tmpdir)
                    with open(tmpdir + ".zip", "rb") as fh: st.download_button("⬇️ Download extracted as zip", fh.read(), file_name="extracted.zip", mime="application/zip")
                    shutil.rmtree(tmpdir)
                st.success("✅ Ekstraksi Selesai (File diekstrak dan dikompres ulang ke ZIP)")
            except Exception: st.error(traceback.format_exc())
    
    elif mode == "Excel -> CSV":
        st.markdown("#### ↔️ Konversi Excel ke CSV")
        file = st.file_uploader("Unggah file Excel:", type=["xlsx"])
        if file and st.button("Proses Konversi ke CSV"):
            # (Logika Excel -> CSV, sama seperti sebelumnya)
            try:
                df = pd.read_excel(file); csv = df.to_csv(index=False).encode("utf-8")
                st.download_button("⬇️ Unduh CSV", csv, "konversi.csv", "text/csv")
                st.success("✅ Konversi berhasil")
            except Exception: st.error(traceback.format_exc())
    
    elif mode == "Word -> PDF (Konversi Teks)":
        st.markdown("#### ↔️ Konversi Word (.docx) ke PDF (Teks Sederhana)")
        file = st.file_uploader("Unggah file Word (.docx):", type=["docx"])
        if file and st.button("Proses Konversi ke PDF"):
            if Document is None: st.error("❌ Error: `python-docx` diperlukan.")
            else:
                # (Logika Word -> PDF, sama seperti sebelumnya)
                try:
                    doc = Document(io.BytesIO(file.read())); text = "\n".join([p.text for p in doc.paragraphs])
                    pdf_buffer = io.BytesIO(); pdf_buffer.write(text.encode("utf-8"))
                    st.download_button("⬇️ Unduh Hasil PDF (raw text)", pdf_buffer.getvalue(), "konversi.pdf", "application/pdf")
                    st.success("✅ Konversi selesai (hanya teks mentah).")
                except Exception: st.error(traceback.format_exc())


# -------------- 5.6 Tentang --------------
if menu == "Tentang":
    add_back_to_dashboard_button() 
    st.subheader("ℹ️ Tentang KAY App – Tools MCU")
    st.markdown("""
    **KAY App** adalah aplikasi serbaguna berbasis Streamlit untuk membantu:
    - **Kompresi & Gambar:** Kompresi, Batch Rename/Format.
    - **PDF:** Gabung, pisah, proteksi, ekstraksi, Reorder/Hapus Halaman, **Batch Rename Sesuai Excel**.
    - **MCU:** Analisis data dan pengorganisasian file.

    ### ⚙️ Kebutuhan Library (Instalasi di Server)

    Untuk mengaktifkan semua fitur, pastikan library berikut terinstal di *environment* Anda:
    - **`PyPDF2`, `pandas`, `openpyxl`, `Pillow`:** (Biasanya sudah terinstal).
    - **`pdfplumber`:** Untuk ekstraksi tabel.
    - **`python-docx`:** Untuk memproses file Word.
    - **`pdf2image` + `poppler`:** Untuk konversi PDF ke Gambar dan Preview gambar.

    Data Anda diproses **secara lokal** di perangkat server tempat Streamlit dijalankan.
    """)

# ----------------- Footer -----------------
st.markdown("""
<hr style="border: none; border-top: 1px solid #cfe2ff; margin-top: 1.5rem; margin-bottom: 0.5rem;">
<div style="text-align:center; color:#5d6d7e; font-size:0.9rem;">
    © 2025 KAY App – Tools MCU | Built with Streamlit 🛠️
</div>

""", unsafe_allow_html=True)
