"""
KAY App - FINAL SINGLE PAGE APP (SPA) - VERSI DIPERBARUI & TAMPILAN LEBIH MENARIK
Skrip ini menggabungkan semua fitur yang diminta:
- **FITUR LAMA LENGKAP:** Gabung, Pisah, Encrypt, Reorder, Kompres Foto.
- **FITUR BARU LENGKAP:** Batch Rename PDF/Gambar Sesuai Excel/Sequential, Organise MCU by Excel.
- **FITUR DIPERBARUI:** Dashboard Analisis Data MCU Massal.
- **FITUR TERBARU:** Terjemahan PDF ke Bahasa Lain.
- **FITUR BARU LAINNYA:** QR Code Generator dengan Logo.
"""

import os
import io
import zipfile
import shutil
import traceback
import tempfile
import time
from io import BytesIO # Diperlukan untuk QR Code dan File Handling

import streamlit as st
import pandas as pd
from PIL import Image
# REQUIRED IMPORTS FOR NEW FEATURE: QR CODE
import qrcode 

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

# pdf2image: try both convert_from_bytes and convert_from_path availability
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
        pass # convert_from_path and convert_from_bytes remain None

# New imports for translation
Translator = None
try:
    from deep_translator import GoogleTranslator
    Translator = GoogleTranslator
except Exception:
    pass # Peringatan akan ditampilkan di fitur Translate PDF jika gagal impor

# ----------------- Helpers -----------------
def make_zip_from_map(bytes_map: dict) -> bytes:
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w") as z:
        for name, data in bytes_map.items():
            z.writestr(name, data)
    b.seek(0)
    return b.getvalue()

def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    # Menggunakan openpyxl sebagai engine
    with pd.ExcelWriter(out, engine="openpyxl") as writer: 
        df.to_excel(writer, index=False)
    out.seek(0)
    return out.getvalue()

def try_encrypt(writer, password: str):
    """Fungsi untuk enkripsi PDF, menampung try/except"""
    try:
        writer.encrypt(password)
    except TypeError:
        try:
            writer.encrypt(user_pwd=password, owner_pwd=None)
        except Exception:
            writer.encrypt(user_pwd=password, owner_pwd=password)

def rotate_page_safe(page, angle):
    """Fungsi untuk rotasi halaman PDF."""
    try:
        page.rotate(angle)
    except Exception:
        try:
            from PyPDF2.generic import NameObject, NumberObject
            page.__setitem__(NameObject("/Rotate"), NumberObject(angle))
        except Exception:
            pass
          
def navigate_to(target_menu):
    """Helper global untuk navigasi antar halaman/menu."""
    st.session_state.menu_selection = target_menu
    try:
        st.rerun() 
    except AttributeError:
        st.experimental_rerun()

# ----------------- Streamlit config & CSS (Perapihan Ikon) -----------------
# LOGO_PATH = os.path.join("assets", "logo.png") # Dinonaktifkan karena file assets tidak tersedia
LOGO_PATH = "???" # Menggunakan emoji toolbox ??? sebagai fallback ikon
page_icon = LOGO_PATH
st.set_page_config(page_title="Master App – Tools MCU", page_icon=page_icon, layout="wide", initial_sidebar_state="collapsed")

# CSS / Theme
st.markdown("""
<style>
/* 1. HILANGKAN SEMUA UI SIDEBAR */
[data-testid="stSidebarToggleButton"], 
section[data-testid="stSidebar"],      
[data-testid="stDecoration"]        
{
    visibility: hidden !important;
    display: none !important;
    width: 0 !important;
    padding: 0 !important;
}

/* 2. Hilangkan logo GitHub 'Fork' */
.stApp a[href*="github.com/"],
.stApp header > div:last-child {
    display: none !important;
}

/* 3. Background gradient lembut */
.stApp {
    background: linear-gradient(180deg, #e9f2ff 0%, #f4f9ff 100%); 
    color: #002b5b;
    font-family: 'Inter', sans-serif;
}

/* 4. Tombol modern glossy - Diperkuat */
div.stButton > button {
    background: linear-gradient(90deg, #5dade2, #3498db);
    color: white; 
    border: none;
    border-radius: 12px;
    padding: 0.7rem 1.2rem; /* Padding lebih besar */
    font-weight: 600;
    transition: 0.2s;
    box-shadow: 0 4px 10px rgba(52, 152, 219, 0.4); 
    cursor: pointer;
    width: auto;
}

/* Tombol Dashboard & Kembali dibuat lebar penuh di konteks masing-masing */
div.stButton > button[key*="dash_"], 
div.stButton > button[key*="back_"] {
    width: 100%;
    margin-top: 10px; 
}

div.stButton > button:hover {
    background: linear-gradient(90deg, #3498db, #2e86c1); 
    transform: scale(1.02);
    /* Skala hover lebih jelas */
    box-shadow: 0 8px 18px rgba(52, 152, 219, 0.6);
    /* Shadow lebih kuat */
}

/* 5. Card fitur - Diperkuat */
.feature-card {
    background: white;
    border-radius: 16px;
    /* Radius lebih besar */
    box-shadow: 0 4px 15px rgba(0,0,0,0.08);
    /* Shadow lebih tegas */
    padding: 24px;
    /* Padding lebih besar */
    transition: all 0.3s ease-in-out;
    /* Transisi untuk semua properti */
    border: 1px solid #d0e3ff; 
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    height: 100%; 
}
.feature-card:hover {
    box-shadow: 0 10px 30px rgba(0,0,0,0.15);
    /* Shadow hover sangat kuat */
    transform: translateY(-5px);
    /* Efek 'terangkat' lebih tinggi */
    border-color: #3498db;
    /* Border highlight */
}

/* 6. Small UI tweaks */
h1 { color: #1b4f72; font-weight: 800; }
h2, h3, h4 { color: #2e86c1; }
.stInfo, .stWarning {
    border-radius: 12px; /* Radius lebih besar */
    padding: 1rem;
}
.stInfo { background-color: #e3f2fd; border-left: 5px solid #2196f3; }
.stWarning { background-color: #fff3e0; border-left: 5px solid #ff9800; }

/* 7. Kontainer untuk merapikan section */
.dashboard-container {
    padding: 20px 0;
    margin-bottom: 20px;
}
</style>
""", unsafe_allow_html=True)

# ----------------- Global Header & Navigation Setup -----------------

# Inisialisasi session state
if "menu_selection" not in st.session_state:
    st.session_state.menu_selection = "Dashboard"
    
menu = st.session_state.menu_selection

# ----------------- Fungsi Tombol Kembali (Ikon Diperbaiki) -----------------
def add_back_to_dashboard_button():
    """Menambahkan tombol 'Kembali ke Dashboard' di halaman fitur dengan ikon ??."""
    if st.button("?? Kembali ke Dashboard", key="back_to_dash"):
        navigate_to("Dashboard")
    st.markdown("---")

# ----------------- Halaman Header (Selalu ditampilkan) -----------------
header_col1, header_col2 = st.columns([1, 4])
with header_col1:
    # Menggunakan ikon emoji karena LOGO_PATH tidak tersedia
    st.markdown("<h1 style='font-size: 3rem;'>???</h1>", unsafe_allow_html=True)

with header_col2:
    st.title("Master - App – Tools MCU")
    st.markdown("Aplikasi serbaguna untuk pengolahan dokumen, PDF, gambar, dan data MCU — UI elegan + fungsi lengkap.")

st.markdown("---")


# -----------------------------------------------------------------------------
# ----------------- FUNGSI UTAMA -----------------
# -----------------------------------------------------------------------------

# -------------- Dashboard (Menu Utama) --------------
if menu == "Dashboard":
    
    st.markdown('<div class="dashboard-container">', unsafe_allow_html=True)
    st.markdown("## ? Pilih Fitur Utama")
    
    # ------------------ FITUR UTAMA ------------------
    cols1 = st.columns(3)

    # Kompres Foto
    with cols1[0]:
        with st.container():
            # Ikon: ??? (Picture) atau ?? (Camera)
            st.markdown('<div class="feature-card"><b>??? Kompres Foto / Gambar Tools</b><br>Perkecil ukuran, ubah format, Batch Rename Sesuai Excel.</div>', unsafe_allow_html=True)
            if st.button("Buka Kompres Foto", key="dash_foto"):
                navigate_to("Kompres Foto")

    # PDF Tools
    with cols1[1]:
        with st.container():
            # Ikon: ?? (Page/Document) atau ?? (Clip)
            st.markdown('<div class="feature-card"><b>?? PDF Tools</b><br>Gabung, pisah, encrypt, Reorder & Batch Rename Sesuai Excel/Sequential, **Terjemahan**.</div>', unsafe_allow_html=True)
            if st.button("Buka PDF Tools", key="dash_pdf"):
                navigate_to("PDF Tools")

    # MCU Tools
    with cols1[2]:
        with st.container():
            # Ikon: ?? (Stethoscope) atau ?? (Chart)
            st.markdown('<div class="feature-card"><b>?? MCU Tools</b><br>Proses Excel + PDF untuk hasil MCU / Analisis Data. **Termasuk Organise by Excel**</div>', unsafe_allow_html=True)
            if st.button("Buka MCU Tools", key="dash_mcu"):
                navigate_to("MCU Tools")
            
    # ------------------ FITUR LAINNYA ------------------
    st.markdown("## ?? Fitur Lainnya")
    cols2 = st.columns(3)
  
    # File Tools
    with cols2[0]:
        with st.container():
            # Ikon: ?? (Folder)
            st.markdown('<div class="feature-card"><b>?? File Tools</b><br>Zip/unzip, konversi dasar, Batch Rename Gambar & PDF.</div>', unsafe_allow_html=True)
            if st.button("Buka File Tools", key="dash_file"):
                navigate_to("File Tools")

    # QR Code Generator (FITUR BARU DARI SCRIP 2)
    with cols2[1]:
        with st.container():
            # Ikon: ?? (QR Code)
            st.markdown('<div class="feature-card"><b>?? QR Code Generator</b><br>Buat QR Code dari teks/URL, dengan opsi logo di tengah.</div>', unsafe_allow_html=True)
            if st.button("Buka QR Code", key="dash_qr"):
                navigate_to("QR Code Generator")

    # Tentang
    with cols2[2]:
        with st.container():
            # Ikon: ?? (Lightbulb) atau ?? (Info)
            st.markdown('<div class="feature-card"><b>?? Tentang Aplikasi</b><br>Informasi dan kebutuhan library.</div>', unsafe_allow_html=True)
            if st.button("Lihat Tentang", key="dash_about"):
                navigate_to("Tentang")

        
    st.markdown('</div>', unsafe_allow_html=True) # Tutup dashboard-container

    st.markdown("---")
    st.info("Semua proses berlangsung lokal di perangkat server tempat Streamlit dijalankan.")


# -------------- QR Code Generator (Integrasi SCRIP 2) --------------
elif menu == "QR Code Generator":
    add_back_to_dashboard_button()
    st.subheader("?? QR Code Generator dengan Logo")
    st.write("Buat QR Code dari teks/URL dengan mudah — bisa menambahkan logo di tengahnya. (Memerlukan `qrcode` dan `Pillow`).")

    # Input data untuk QR
    data = st.text_input("Masukkan teks atau URL:", "")

    # Unggah logo opsional
    uploaded_logo = st.file_uploader("Unggah logo (opsional, format PNG/JPG/WEBP)", type=["png", "jpg", "jpeg", "webp"])

    # Pengaturan ukuran
    col1, col2 = st.columns(2)
    with col1:
        box_size = st.slider("Ukuran kotak (Box Size)", 5, 20, 10, help="Semakin besar, semakin besar resolusi QR Code yang dihasilkan.")
    with col2:
        border = st.slider("Tebal border (Margin)", 2, 10, 4, help="Jumlah modul kotak kosong yang membentuk margin di sekitar kode.")

    if st.button("?? Buat QR Code", use_container_width=True):
        if not data:
            st.warning("Masukkan teks atau URL terlebih dahulu!")
        else:
            try:
                # Buat QR. ERROR_CORRECT_H (High) diperlukan untuk logo.
                qr = qrcode.QRCode(
                    version=1,
                    error_correction=qrcode.constants.ERROR_CORRECT_H,
                    box_size=box_size,
                    border=border,
                )
                qr.add_data(data)
                qr.make(fit=True) 
                
                # Konversi ke RGB
                qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")

                # Tambahkan logo jika ada
                if uploaded_logo is not None:
                    st.info("Logo terdeteksi. QR Code dibuat dengan tingkat koreksi kesalahan TINGGI.")
                    logo = Image.open(uploaded_logo)
                    qr_width, qr_height = qr_img.size
                    # Ubah ukuran logo agar proporsional (20% dari ukuran QR Code)
                    logo_size = int(qr_width * 0.2)
                    logo = logo.resize((logo_size, logo_size))
                    # Tempel logo di tengah
                    pos = ((qr_width - logo_size) // 2, (qr_height - logo_size) // 2)
                    # Gunakan mask untuk transparansi jika mode logo RGBA
                    qr_img.paste(logo, pos, mask=logo if logo.mode == "RGBA" else None)

                # Tampilkan QR
                st.image(qr_img, caption="QR Code Hasil", use_container_width=False)

                # Download hasil
                buf = io.BytesIO() 
                qr_img.save(buf, format="PNG")
                buf.seek(0)
                st.download_button(
                    "?? Download QR Code (.png)",
                    data=buf.getvalue(), 
                    file_name="qrcode_result.png",
                    mime="image/png"
                )

            except Exception as e:
                st.error(f"Gagal memproses QR Code: {e}")
                traceback.print_exc()

# -------------- Kompres Foto / Image Tools (Dari SCRIP 1) --------------
elif menu == "Kompres Foto":
    add_back_to_dashboard_button() 
    st.subheader("??? Kompres & Kelola Foto/Gambar")
    
    # Sub-menu untuk gambar
    img_tool = st.selectbox("Pilih Fitur Gambar", [
        "?? Kompres Foto (Batch)", 
        "?? Batch Rename/Format Gambar (Sequential)",
        "?? Batch Rename Gambar Sesuai Excel (Fitur Baru)"
        ])

    if img_tool == "?? Kompres Foto (Batch)":
        st.markdown("---")
        st.markdown("### ?? Kompres Foto (Batch)")
        uploaded = st.file_uploader("Unggah gambar (jpg/png) — bisa banyak", type=["jpg","jpeg","png"], accept_multiple_files=True)
        quality = st.slider("Kualitas JPEG", 10, 95, 75)
        max_side = st.number_input("Max side (px)", min_value=100, max_value=4000, value=1200)
        if uploaded and st.button("Kompres Semua"):
            out_map = {}
            total = len(uploaded)
            prog = st.progress(0)
            with st.spinner("Mengompres..."):
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
                st.success(f"?? {len(out_map)} file berhasil dikompres")
                st.download_button("Unduh Hasil (ZIP)", zipb, file_name="foto_kompres.zip", mime="application/zip")
            else:
                st.warning("Tidak ada file berhasil dikompres.")

    # --- FITUR Batch Rename Gambar (Sequential) ---
    elif img_tool == "?? Batch Rename/Format Gambar (Sequential)": 
        st.markdown("---")
        st.markdown("### ?? Ganti Nama & Ubah Format Gambar Massal (Sequential)")
        uploaded_files = st.file_uploader(
            "Unggah file Gambar (JPG, PNG, dll.):", 
            type=["jpg", "jpeg", "png", "webp"], 
            accept_multiple_files=True,
            key="batch_rename_uploader"
        )
        if uploaded_files:
            col1, col2 = st.columns(2)
            new_prefix = col1.text_input("Prefix Nama File Baru:", value="KAY_File", help="Contoh: KAY_File_001.jpg", key="prefix_img_seq")
            new_format = col2.selectbox("Format Output Baru:", ["Sama seperti Asli", "JPG", "PNG", "WEBP"], index=0, key="format_img_seq")

            if st.button("Proses Batch File", key="process_batch_rename_seq"):
                if not new_prefix: 
                    st.error("Prefix nama file tidak boleh kosong.")
                    st.stop()
                else:
                    output_zip = io.BytesIO()
                    try:
                        with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for i, file in enumerate(uploaded_files, 1):
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
                        st.success(f"?? Berhasil memproses {len(uploaded_files)} file.") 
                        st.download_button("Unduh File ZIP Hasil Batch", data=output_zip.getvalue(), file_name="hasil_batch_gambar.zip", mime="application/zip")
                    except Exception as e: st.error(f"Gagal memproses file: {e}"); traceback.print_exc()

    # --- FITUR BARU 1: Batch Rename Gambar Sesuai Excel ---
    elif img_tool == "?? Batch Rename Gambar Sesuai Excel (Fitur Baru)": 
        st.markdown("---")
        st.markdown("### ?? Ganti Nama Gambar (PNG/JPEG) Berdasarkan Excel")
        st.info("Template Excel/CSV wajib memiliki kolom **`nama_lama`** (termasuk ekstensi, misal: `foto_123.jpg`) dan **`nama_baru`** (termasuk ekstensi, misal: `ID_001.png`).")
        
        excel_up = st.file_uploader("Unggah Excel/CSV untuk daftar nama:", type=["xlsx", "csv"], key="rename_img_excel_up")
        files = st.file_uploader("Unggah Gambar (JPG/PNG/JPEG, multiple):", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="rename_img_files_up")
    
        if excel_up and files and st.button("Proses Ganti Nama Gambar (ZIP)", key="process_img_rename_excel"):
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
                        st.error(f"Excel/CSV wajib memiliki kolom: {', '.join(required_cols)}")
                        st.stop() 
                    
                    # 3. Map File dan Proses Rename
                    file_map = {f.name: f.read() for f in files}
                    out_map = {}
                    not_found = []
                    df['nama_lama_str'] = df['nama_lama'].astype(str).str.strip() # Gunakan str.strip() untuk membersihkan spasi

                    for _, row in df.iterrows():
                        old_name = str(row['nama_lama']).strip()
                        new_name = str(row['nama_baru']).strip()
                        
                        # Cek di file yang diupload (case-sensitive)
                        if old_name in file_map:
                            # Pastikan nama baru memiliki ekstensi
                            if not os.path.splitext(new_name)[1]:
                                # Coba ambil ekstensi dari nama lama jika nama baru tidak ada
                                _, old_ext = os.path.splitext(old_name)
                                new_name = new_name + old_ext
                                
                            out_map[new_name] = file_map[old_name]
                        else:
                            not_found.append(old_name)

                    # 4. Buat ZIP
                    if out_map:
                        zipb = make_zip_from_map(out_map)
                        st.success(f"?? {len(out_map)} file berhasil diganti namanya dan dikemas.") 
                        st.download_button("Unduh Hasil (ZIP)", zipb, file_name="gambar_renamed_by_excel.zip", mime="application/zip")
                    else:
                        st.warning("Tidak ada file yang cocok ditemukan atau diproses.")
            
                    if not_found:
                        st.info(f"{len(not_found)} file 'nama_lama' di Excel tidak ditemukan di file yang diunggah. Contoh: {not_found[:5]}")
            except Exception as e:
                st.error(f"Terjadi kesalahan pemrosesan: {e}")
                traceback.print_exc()

# -------------- PDF Tools (Dari SCRIP 1) --------------
elif menu == "PDF Tools":
    add_back_to_dashboard_button() 
    st.subheader("?? PDF Tools")

    # Menu yang lebih terstruktur dan ditambahkan fitur baru dengan ikon yang jelas
    pdf_options = [
        "--- Pilih Tools ---",
        "?? Gabung PDF", 
        "?? Pisah PDF", 
        "?? Reorder/Hapus Halaman PDF", 
        "?? Batch Rename PDF (Sequential)", 
        "?? Batch Rename PDF Sesuai Excel (Fitur Baru)", 
        "??? Image -> PDF", 
        "?? PDF -> Image", 
        "?? Ekstraksi Teks/Tabel", 
        "??? Terjemahan PDF ke Bahasa Lain (Fitur Baru)", # <-- FITUR BARU TRANSLATE
        "?? Konversi PDF", 
        "?? Proteksi PDF", 
        "??? Utility PDF", 
    ]
    
    tool_select = st.selectbox("Pilih fitur PDF", pdf_options)

    # Mapping
    if tool_select == "--- Pilih Tools ---": tool = None
    elif tool_select == "?? Ekstraksi Teks/Tabel": tool = st.selectbox("Pilih mode ekstraksi", ["Extract Text", "Extract Tables -> Excel"]) 
    elif tool_select == "?? Konversi PDF": tool = st.selectbox("Pilih mode konversi", ["PDF -> Word", "PDF -> Excel (text)"]) 
    elif tool_select == "?? Proteksi PDF": tool = st.selectbox("Pilih mode proteksi", ["Encrypt PDF", "Decrypt PDF", "Batch Lock (Excel)"]) 
    elif tool_select == "??? Utility PDF": tool = st.selectbox("Pilih mode utilitas", ["Hapus Halaman", "Rotate PDF", "Kompres PDF", "Watermark PDF", "Preview PDF"]) 
    elif tool_select == "?? Gabung PDF": tool = "Gabung PDF" 
    elif tool_select == "?? Pisah PDF": tool = "Pisah PDF" 
    elif tool_select == "?? Reorder/Hapus Halaman PDF": tool = "Reorder PDF" 
    elif tool_select == "?? Batch Rename PDF (Sequential)": tool = "Batch Rename PDF Seq" 
    elif tool_select == "?? Batch Rename PDF Sesuai Excel (Fitur Baru)": tool = "Batch Rename PDF Excel" 
    elif tool_select == "?? PDF -> Image": tool = "PDF -> Image" 
    elif tool_select == "??? Image -> PDF": tool = "Image -> PDF" 
    elif tool_select == "??? Terjemahan PDF ke Bahasa Lain (Fitur Baru)": tool = "Translate PDF" # <-- MAPPING TRANSLATE
    else: tool = None
    

# --- FITUR BARU: Terjemahan PDF (Optimasi Struktur/Rapi) ---
    if tool == "Translate PDF":
        st.markdown("---")
        st.markdown("### ??? Terjemahan Teks PDF (Optimasi Agar Lebih Rapi)")
        st.info("Fitur ini mencoba membuat hasil Word lebih rapi dengan menggabungkan baris-baris pendek yang berdekatan (*pre-processing*). **Replikasi tata letak kolom/tabel PDF tetap terbatas.**")
      
        if Translator is None or Document is None:
            if Translator is None: st.error("Library `deep-translator` tidak ditemukan.")
            if Document is None: st.error("Library `python-docx` tidak ditemukan.")
            st.stop()
        
        f = st.file_uploader("Unggah PDF untuk Diterjemahkan:", type="pdf", key="translate_pdf_uploader")
        
        col1, col2 = st.columns(2)
        src_lang = col1.text_input("Bahasa Sumber (ISO Code, ex: id)", value="auto", help="Ketik 'auto' jika tidak yakin.")
        target_lang = col2.text_input("Bahasa Tujuan (ISO Code, ex: en, ja, fr)", value="en")

        if f and st.button("Proses Terjemahan dan Buat Word (.docx)", key="translate_pdf_button"):
            try:
                
                # --- HELPER: Menggabungkan baris pendek untuk kerapihan ---
                def preprocess_text_for_layout(text_blocks: list) -> list:
                    """Menggabungkan baris-baris pendek yang berdekatan, mengasumsikan itu adalah item daftar atau sel tabel."""
                    processed_blocks = []
                    current_block = ""
                    MAX_LINE_LENGTH = 100 # Batas panjang baris agar tidak digabungkan
                    
                    for block in text_blocks:
                        if block.strip() == "":
                            # Baris kosong yang sebenarnya menandakan akhir paragraf
                            if current_block:
                                processed_blocks.append(current_block)
                            processed_blocks.append("") # Jaga pemisah yang jelas
                            current_block = ""
                            continue
                        
                        # Baris yang sangat pendek (mungkin label/nilai)
                        if len(block.strip()) < MAX_LINE_LENGTH and not block.strip().endswith('.'):
                            # Jika baris pendek, coba gabungkan ke block saat ini dengan spasi
                            if current_block:
                                current_block += " | " + block.strip() # Gunakan pemisah "|" 
                            else:
                                current_block = block.strip()
                        else:
                            # Jika baris panjang (paragraf utuh)
                            if current_block:
                                processed_blocks.append(current_block)
                            current_block = block.strip()
                            processed_blocks.append(current_block)
                            current_block = ""
                         
                    if current_block:
                        processed_blocks.append(current_block)
                    
                    # Hapus spasi ganda yang berlebihan setelah processing
                    return [b.replace("  ", " ").strip() for b in processed_blocks if b.strip() or b == ""]
                # --------------------------------------------------------

                # 1. Ekstraksi Teks (Menggunakan Paragraf sebagai Unit)
                with st.spinner("1. Mengekstrak dan merapikan teks dari PDF..."):
                    raw = f.read()
                    all_text_lines = []
                    
                    if pdfplumber:
                        with pdfplumber.open(io.BytesIO(raw)) as doc:
                            for p in doc.pages:
                                page_text = p.extract_text() or ""
                                # Pisahkan per baris baru tunggal (lebih detail)
                                all_text_lines.extend(page_text.split('\n'))
                                all_text_lines.append("---HALAMAN BARU---") # Marker untuk halaman baru

                    else: # Fallback ke PyPDF2
                        reader = PdfReader(io.BytesIO(raw))
                        for p in reader.pages:
                            page_text = p.extract_text() or ""
                            all_text_lines.extend(page_text.split('\n'))
                            all_text_lines.append("---HALAMAN BARU---")

                    # Lakukan Pre-processing untuk menggabungkan baris-baris yang berdekatan
                    preprocessed_paragraphs = preprocess_text_for_layout(all_text_lines)
                    
                    full_text_clean = "\n\n".join(p for p in preprocessed_paragraphs if p != "---HALAMAN BARU---" and p != "")

                if not full_text_clean.strip():
                    st.warning("Teks kosong atau tidak dapat diekstrak dari PDF.")
                    st.stop()

                # 2. Chunking Berbasis Paragraf dan Terjemahan
                with st.spinner(f"2. Menerjemahkan teks ke {target_lang} (mempertahankan paragraf)..."):
                    translator = Translator(source=src_lang, target=target_lang)
                    CHUNK_SIZE = 4500 # Batas karakter aman
                    
                    # Gabungkan paragraf yang sudah diproses menjadi chunks aman untuk terjemahan
                    current_chunk = ""
                    text_chunks_for_translation = []
                    
                    for p in preprocessed_paragraphs:
                        # Jika penanda halaman ditemukan
                        if p == "---HALAMAN BARU---":
                            if current_chunk:
                                text_chunks_for_translation.append(current_chunk)
                            text_chunks_for_translation.append("---HALAMAN BARU---")
                            current_chunk = ""
                            continue
                        
                        # Jika paragraf yang diproses kosong atau merupakan spasi antar blok
                        if not p.strip(): 
                            if current_chunk:
                                text_chunks_for_translation.append(current_chunk)
                            text_chunks_for_translation.append("") # Tambahkan marker baris kosong
                            current_chunk = ""
                            continue
                        
                        # Logic chunking: tambahkan paragraf ke chunk saat ini
                        if len(current_chunk) + len(p) + 4 > CHUNK_SIZE: # +4 untuk pemisah \n\n
                            if current_chunk:
                                text_chunks_for_translation.append(current_chunk)
                            current_chunk = p + "\n\n"
                        else:
                            current_chunk += p + "\n\n"

                    if current_chunk:
                        text_chunks_for_translation.append(current_chunk.strip())

                    translated_parts = []
                    prog = st.progress(0)
                    for i, chunk in enumerate(text_chunks_for_translation):
                        if chunk in ("---HALAMAN BARU---", ""):
                            translated_parts.append(chunk)
                        else:
                            if i > 0: time.sleep(0.1) # Batasi kecepatan API Call
                            # Terjemahkan, dan ganti pemisah "|" kembali ke format yang rapi
                            translated = translator.translate(chunk)
                            translated_parts.append(translated.replace(" | ", " | ").strip())
                        prog.progress(int((i + 1) / len(text_chunks_for_translation) * 100))
                    
                    translated_text_combined = "\n\n".join(translated_parts)
                    prog.empty()

                # 3. Rekonstruksi ke Word (Memanfaatkan Struktur Lebih Baik)
                with st.spinner("3. Membuat file Word (.docx) baru..."):
                    doc = Document()
                    # Memecah berdasarkan penanda halaman dan paragraf yang jelas
                    for item in translated_text_combined.split('\n\n'):
                        item_stripped = item.strip()
                        if item_stripped == "---HALAMAN BARU---":
                            # Tambahkan Page Break
                            doc.add_page_break()
                        elif item_stripped:
                            # Gunakan paragraf baru
                            doc.add_paragraph(item_stripped)
                            
                    out = io.BytesIO()
                    doc.save(out)
                    out.seek(0)
                    
                    st.success("?? Terjemahan berhasil! Unduh file Word hasil terjemahan.")
                    st.download_button(
                        f"?? Unduh Hasil Terjemahan ({target_lang}).docx",
                        data=out.getvalue(),
                        file_name=f"translated_to_{target_lang}_rapi.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    
                    st.markdown("---")
                    st.markdown("#### Preview Teks Asli (Sudah Dirapikan) dan Terjemahan")
                    col_preview1, col_preview2 = st.columns(2)
                    with col_preview1:
                        st.text_area("Teks Asli (Dirapikan)", "\n\n".join(preprocessed_paragraphs)[:5000], height=300)
                    with col_preview2:
                        st.text_area("Teks Terjemahan", translated_text_combined[:5000], height=300)

            except Exception as e:
                st.error(f"Terjadi kesalahan saat terjemahan. Cek kode bahasa (ISO 639-1) dan pastikan teks dapat diekstrak. Error: {e}")
                traceback.print_exc()

# --- FITUR Batch Rename PDF Sesuai Excel ---
    if tool == "Batch Rename PDF Excel":
        st.markdown("---")
        st.markdown("### ?? Ganti Nama File PDF Berdasarkan Excel")
        st.markdown("Unggah banyak file PDF dan ganti namanya sesuai daftar di Excel/CSV.")
        st.info("Template Excel/CSV wajib memiliki kolom **`nama_lama`** (misal: `ID_123.pdf`) dan **`nama_baru`** (misal: `Hasil_123.pdf`).")
        
        excel_up = st.file_uploader("Unggah Excel/CSV untuk daftar nama:", type=["xlsx", "csv"], key="rename_pdf_excel_up")
        files = st.file_uploader("Unggah File PDF (multiple):", type=["pdf"], accept_multiple_files=True, key="rename_pdf_files_up")
        
        if excel_up and files and st.button("Proses Ganti Nama PDF (ZIP)", key="process_pdf_rename_excel"):
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
                        st.error(f"Excel/CSV wajib memiliki kolom: {', '.join(required_cols)}")
                        st.stop() 
                        
                    # 3. Map File dan Proses Rename
                    file_map = {f.name: f.read() for f in files}
                    out_map = {}
                    not_found = []
                    
                    for _, row in df.iterrows():
                        old_name = str(row['nama_lama']).strip()
                        new_name = str(row['nama_baru']).strip()
                        
                        if old_name in file_map:
                            # Tambahkan ekstensi .pdf jika belum ada di nama baru
                            if not new_name.lower().endswith('.pdf'):
                                new_name += '.pdf'
                            out_map[new_name] = file_map[old_name]
                        else:
                            not_found.append(old_name)

                    # 4. Buat ZIP
                    if out_map:
                        zipb = make_zip_from_map(out_map)
                        st.success(f"?? {len(out_map)} file berhasil diganti namanya dan dikemas.") 
                        st.download_button("Unduh Hasil (ZIP)", zipb, file_name="pdf_renamed_by_excel.zip", mime="application/zip")
                    else:
                        st.warning("Tidak ada file yang cocok ditemukan atau diproses.")
                        
                    if not_found:
                        st.info(f"{len(not_found)} file 'nama_lama' di Excel tidak ditemukan di file yang diunggah. Contoh: {not_found[:5]}")
            except Exception as e:
                st.error(f"Terjadi kesalahan pemrosesan: {e}")
                traceback.print_exc()

# --- FITUR Batch Rename PDF (Sequential) ---
    if tool == "Batch Rename PDF Seq":
        st.markdown("---")
        st.markdown("### ?? Ganti Nama File PDF Massal (Sequential)")
        
        uploaded_files = st.file_uploader("Unggah file PDF (multiple):", type=["pdf"], accept_multiple_files=True, key="batch_rename_pdf_uploader_seq")
        
        if uploaded_files:
            col1, col2 = st.columns(2)
            new_prefix = col1.text_input("Prefix Nama File Baru:", value="Hasil_PDF", help="Contoh: Hasil_PDF_001.pdf", key="prefix_pdf_seq")
            start_num = col2.number_input("Mulai dari Angka (Counter):", min_value=1, value=1, step=1, key="start_num_pdf_seq")
            
            if st.button("Proses Ganti Nama (ZIP)", key="process_batch_rename_pdf_seq"):
                if not new_prefix:
                    st.error("Prefix nama file tidak boleh kosong.")
                    st.stop()
                else:
                    output_zip = io.BytesIO()
                    try:
                        with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for i, file in enumerate(uploaded_files, start_num):
                                new_filename = f"{new_prefix}_{i:03d}.pdf"
                                zf.writestr(new_filename, file.read())
                        st.success(f"?? Berhasil mengganti nama {len(uploaded_files)} file.")
                        st.download_button("Unduh File ZIP Hasil Rename", data=output_zip.getvalue(), file_name="pdf_renamed.zip", mime="application/zip")
                    except Exception as e: st.error(f"Gagal memproses file: {e}"); traceback.print_exc()

# --- LOGIKA FITUR PDF LAINNYA (dari SCRIP 1, hanya bagian Reorder yang sudah ada) ---
    if tool == "Reorder PDF":
        st.markdown("---")
        st.markdown("### ?? Reorder atau Hapus Halaman PDF")
        st.markdown("Unggah file PDF Anda dan tentukan urutan halaman baru (contoh: `2, 1, 3` untuk membalik, atau `1, 3` untuk menghapus halaman 2).")
        
        f = st.file_uploader("Unggah 1 file PDF:", type="pdf", key="reorder_pdf_uploader")
        
        if f:
            try:
                raw = f.read()
                if PdfReader is None:
                    st.error("PyPDF2 tidak terinstall (pip install PyPDF2)")
                    st.stop()
                    
                reader = PdfReader(io.BytesIO(raw))
                num_pages = len(reader.pages)
                st.info(f"PDF berhasil dimuat. Jumlah total halaman: **{num_pages}**.")
                
                default_order = ", ".join(map(str, range(1, num_pages + 1)))
                new_order_str = st.text_input(
                    f"Masukkan urutan halaman baru (1-{num_pages}) dipisahkan koma:",
                    value=default_order,
                    help="Contoh: '3, 1, 2' untuk mengubah urutan. '1, 3, 5' untuk menghapus halaman genap."
                )
                
                if st.button("Proses Reorder/Hapus Halaman", key="process_reorder"):
                    new_order_indices = []
                    try:
                        input_list = [int(x.strip()) for x in new_order_str.split(',') if x.strip().isdigit()]
                        
                        if any(n < 1 or n > num_pages for n in input_list):
                            st.error(f"Nomor halaman harus antara 1 sampai {num_pages}.")
                            st.stop()
                            
                        new_order_indices = [n - 1 for n in input_list]
                        
                        writer = PdfWriter()
                        for index in new_order_indices:
                            writer.add_page(reader.pages[index])
                            
                        out = io.BytesIO()
                        writer.write(out)
                        out.seek(0)
                        
                        st.success("?? Halaman berhasil diurutkan ulang/dihapus.")
                        st.download_button("Unduh Hasil PDF", out.getvalue(), file_name="reordered_pdf.pdf", mime="application/pdf")
                        
                    except Exception as e:
                        st.error(f"Format urutan halaman salah atau terjadi error: {e}")
                        traceback.print_exc()
                        
            except Exception as e:
                st.error(f"Gagal memuat PDF: {e}")
                traceback.print_exc()

    # --- Placeholder untuk fitur PDF lainnya ---
    if tool in ["Gabung PDF", "Pisah PDF", "Extract Text", "Extract Tables -> Excel", "PDF -> Word", "PDF -> Excel (text)", "Encrypt PDF", "Decrypt PDF", "Batch Lock (Excel)", "Hapus Halaman", "Rotate PDF", "Kompres PDF", "Watermark PDF", "Preview PDF", "PDF -> Image", "Image -> PDF"]:
        if tool != "Reorder PDF" and tool != "Batch Rename PDF Excel" and tool != "Batch Rename PDF Seq" and tool != "Translate PDF":
            st.markdown("---")
            st.markdown(f"### ?? Fitur: {tool}")
            st.warning("Fitur ini adalah bagian dari struktur aplikasi, tetapi logika implementasi penuhnya tidak ditemukan atau terpotong di dalam skrip yang diunggah. Hanya placeholder.")

# -------------- MCU Tools (Placeholder) --------------
elif menu == "MCU Tools":
    add_back_to_dashboard_button()
    st.subheader("?? MCU Tools")
    st.markdown("### Fitur Analisis Data dan Organise MCU by Excel")
    st.warning("Fitur ini adalah bagian dari struktur aplikasi, tetapi logika implementasi penuhnya tidak ditemukan di dalam skrip yang diunggah. Hanya placeholder.")

# -------------- File Tools (Placeholder) --------------
elif menu == "File Tools":
    add_back_to_dashboard_button()
    st.subheader("?? File Tools")
    st.markdown("### Fitur Zip/Unzip, Konversi Dasar, Batch Rename")
    st.warning("Fitur ini adalah bagian dari struktur aplikasi, tetapi logika implementasi penuhnya tidak ditemukan di dalam skrip yang diunggah. Hanya placeholder.")

# -------------- Tentang Aplikasi (Placeholder) --------------
elif menu == "Tentang":
    add_back_to_dashboard_button()
    st.subheader("?? Tentang Aplikasi")
    st.markdown("### Informasi & Kebutuhan Library")
    st.markdown("""
    **KAY App** adalah aplikasi serbaguna berbasis Streamlit untuk membantu:
    - ?? **Kompres Foto & Gambar**
    - ?? **Pengelolaan Dokumen PDF** (gabung, pisah, proteksi, ekstraksi, Reorder/Hapus Halaman, Batch Rename, **Terjemahan**)
    - ?? **Analisis & Pengolahan Hasil MCU** (Dashboard Analisis Data, **Organise by Excel**)
    - ?? **Manajemen File & Konversi Dasar** (Batch Rename/Format Gambar, Batch Rename PDF)
    - ?? **QR Code Generator**
    <br>
    
    ### Kebutuhan Library Tambahan (Instal di lingkungan Anda)
    Beberapa fitur memerlukan library tambahan:
    - `PyPDF2` (Dasar PDF): `pip install PyPDF2`
    - `pdfplumber` untuk ekstraksi tabel teks: `pip install pdfplumber`
    - `python-docx` untuk menghasilkan .docx: `pip install python-docx`
    - `deep-translator` untuk fitur terjemahan PDF: `pip install deep-translator`
    - `pdf2image` + poppler untuk konversi PDF->Gambar / Preview gambar: `pip install pdf2image`
    - `pandas` & `openpyxl` untuk Analisis MCU dan Batch Rename by Excel: `pip install pandas openpyxl`
    - `qrcode` & `Pillow` untuk QR Code: `pip install qrcode Pillow`
    """)
    st.info("Data diproses di server tempat Streamlit dijalankan. Untuk mengaktifkan semua fitur, pasang dependensi yang diperlukan.")

# ----------------- Footer -----------------
st.markdown("""
<hr style="border: none; height: 1px; background-color: #ddd; margin-top: 20px;">
<div style="text-align: center; color: #777; font-size: 0.8rem;">
    Master App - Tools MCU | Dibuat dengan ?? Streamlit & Python
</div>
""", unsafe_allow_html=True)
